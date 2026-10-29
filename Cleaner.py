# Python
"""
Spreadsheet data cleaning workflows for customer and invoice data.

The module reads CSV and Excel files, normalizes recognized fields, and
separates clean rows from rows that need human review or rejection. It does
not connect to or modify a live database.
"""

# Standard-library imports used for parsing, logging, command-line options, and
# file paths.
import logging
import argparse
import math
import re
import sys
from pathlib import Path
from datetime import date, datetime, timedelta
from dataclasses import dataclass, field
from typing import Iterable, List, Dict, Any, Tuple, Optional

# Third-party libraries used for spreadsheet processing and data validation.
import pandas as pd
from thefuzz import fuzz
from email_validator import validate_email, EmailNotValidError
import phonenumbers
from phonenumbers import PhoneNumberFormat
import xlwt

# Configure application logging for errors encountered by the command-line run.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


@dataclass
class CleaningReport:
    """Keep the counts and row-level findings from one cleaning run.

    The same report object is used by both customer and invoice workflows so
    the command line and exported reports can summarize either input type.
    """
    original_file: str
    total_rows: int = 0
    processed_rows: int = 0
    cleaned_count: int = 0
    review_count: int = 0
    rejected_count: int = 0
    duplicates_removed: int = 0
    fuzzy_matches: List[Dict[str, Any]] = field(default_factory=list)
    potential_matches_limited: bool = False
    rejected_rows: List[Dict[str, Any]] = field(default_factory=list)

    def summary_stats(self) -> Dict[str, float]:
        """Return outcome percentages, or an empty mapping for an empty file."""
        if self.total_rows == 0:
            return {}
        return {
            "Clean Rate": (self.cleaned_count / self.total_rows) * 100,
            "Review Rate": (self.review_count / self.total_rows) * 100,
            "Rejection Rate": (self.rejected_count / self.total_rows) * 100,
            "Duplicate Rate": (self.duplicates_removed / self.total_rows) * 100,
        }


# Normalization helpers convert one cell at a time. They return a normalized
# value plus validity information so the row workflows can decide whether to
# keep a row, mark it for review, or reject it.
def normalize_name(name: Any) -> Optional[str]:
    """Trim a name, collapse whitespace, and apply title case.

    Return None for a blank cell so callers can distinguish missing names.
    """
    if pd.isna(name) or str(name).strip() == "":
        return None
    name = re.sub(r'\s+', ' ', str(name).strip())
    return name.title()


def get_comparison_name(name: Any) -> Optional[str]:
    """Create a comparison-only name without changing the exported name."""
    if pd.isna(name) or str(name).strip() == "":
        return None
    return re.sub(r'\s+', ' ', str(name).strip()).lower()


def normalize_email(email: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """Return (normalized address, valid, error) for an email cell.

    Deliverability is intentionally not checked: this validates address
    syntax without making a network request or claiming the mailbox exists.
    """
    if not _has_value(email):
        return None, True, None # Missing is acceptable but flagged later

    email_str = str(email).strip().lower()
    try:
        # Validate the address syntax without requiring a live DNS lookup.
        valid = validate_email(email_str, check_deliverability=False)
        return valid.normalized, True, None
    except EmailNotValidError as e:
        return None, False, str(e)


def normalize_phone(phone: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """Return (E.164 number, valid, error), assuming US for unprefixed input.

    Numeric spreadsheet cells are accepted only when they represent whole,
    finite numbers; the returned phone value is a string to preserve its
    international prefix.
    """
    if not _has_value(phone):
        return None, True, None

    if isinstance(phone, (int, float)) and not isinstance(phone, bool):
        if isinstance(phone, float):
            if not math.isfinite(phone) or not phone.is_integer():
                return None, False, "Phone number must be a whole finite number"
            phone = int(phone)
    phone_str = str(phone).strip()
    try:
        # Parses using Google's libphonenumber. Assumes US if no prefix.
        parsed = phonenumbers.parse(phone_str, "US")
        if not phonenumbers.is_valid_number(parsed):
            return None, False, "Invalid phone number format"

        # Standardize to E.164 format (+15551234567)
        return phonenumbers.format_number(parsed, PhoneNumberFormat.E164), True, None
    except phonenumbers.NumberParseException as e:
        return None, False, str(e)


def _has_value(value: Any) -> bool:
    """Safely test whether a spreadsheet cell contains a nonblank value."""
    if value is None or pd.isna(value):
        return False
    return str(value).strip() != ""


def normalize_date(date_val: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """Return an ISO date, validity flag, and error for a supported date.

    Blank cells are allowed. Excel serial numbers use Excel's 1899-12-30
    epoch, while slash-formatted strings are rejected if both day-first and
    month-first interpretations are possible and differ.
    """
    if not _has_value(date_val):
        return None, True, None
    if isinstance(date_val, (date, datetime)):
        return date_val.strftime("%Y-%m-%d"), True, None
    if isinstance(date_val, (int, float)) and not isinstance(date_val, bool):
        if not math.isfinite(date_val) or not 1 <= date_val <= 2_958_465:
            return None, False, "Invalid Excel date serial"
        parsed_date = datetime(1899, 12, 30) + timedelta(days=float(date_val))
        return parsed_date.strftime("%Y-%m-%d"), True, None

    date_str = str(date_val).strip()
    # Slash-separated dates can be valid in both month/day and day/month order.
    if re.fullmatch(r"\d{1,2}/\d{1,2}/\d{4}", date_str):
        month_first = day_first = None
        for fmt, destination in (
            ("%m/%d/%Y", "month_first"),
            ("%d/%m/%Y", "day_first"),
        ):
            try:
                parsed = datetime.strptime(date_str, fmt)
            except ValueError:
                continue
            if destination == "month_first":
                month_first = parsed
            else:
                day_first = parsed
        if month_first and day_first and month_first != day_first:
            return None, False, "Ambiguous date format"
        parsed = month_first or day_first
        if parsed:
            return parsed.strftime("%Y-%m-%d"), True, None
        return None, False, "Invalid date format"

    # These formats have an unambiguous year-first or named-month ordering.
    formats = ["%Y-%m-%d", "%Y/%m/%d", "%b %d, %Y", "%B %d, %Y"]

    for fmt in formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d"), True, None
        except ValueError:
            continue
    return None, False, "Ambiguous or invalid date format"


def normalize_currency(val: Any) -> Tuple[Optional[float], bool, Optional[str]]:
    """Parse a currency-like cell as (number, valid, error).

    Parentheses are accepted as accounting notation for negative amounts.
    Empty values are allowed; malformed grouping and non-finite numbers are
    rejected instead of silently converted.
    """
    if not _has_value(val):
        return None, True, None

    s = str(val).strip()
    parenthesized = s.startswith("(") and s.endswith(")")
    if parenthesized:
        s = s[1:-1].strip()
    elif "(" in s or ")" in s:
        return None, False, "Invalid numeric currency value"

    if s.count("$") > 1 or ("$" in s and not re.match(r"^[+-]?\$", s)):
        return None, False, "Invalid numeric currency value"
    s = re.sub(r"^([+-]?)\$\s*", r"\1", s).strip()
    if parenthesized and s.startswith(("+", "-")):
        return None, False, "Invalid numeric currency value"

    # Validate thousands separators before removing them; float() alone accepts
    # malformed values such as "1,2,3" after commas are stripped.
    number_pattern = r"(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)"
    if not re.fullmatch(rf"[+-]?{number_pattern}(?:[eE][+-]?\d+)?", s):
        return None, False, "Invalid numeric currency value"
    if parenthesized:
        s = f"-{s}"
    try:
        number = float(s.replace(",", ""))
        if not math.isfinite(number):
            raise ValueError("Currency value must be finite")
        return number, True, None
    except ValueError:
        return None, False, "Invalid numeric currency value"


def _normalize_invoice_date(value: Any) -> Tuple[Optional[str], bool, Optional[str]]:
    """Normalize invoice dates, treating slash/hyphen numeric strings as US dates.

    Invoice-specific formats are tried before the general date parser because
    invoice exports commonly use month-first dates and YYYY.MM.DD values.
    """
    if isinstance(value, str):
        date_text = value.strip()
        for date_format in ("%m/%d/%Y", "%m-%d-%Y"):
            try:
                parsed = datetime.strptime(date_text, date_format)
                return parsed.strftime("%Y-%m-%d"), True, None
            except ValueError:
                continue

    normalized, valid, error = normalize_date(value)
    if valid or not _has_value(value) or not isinstance(value, str):
        return normalized, valid, error

    date_text = value.strip()
    formats = ("%Y.%m.%d",)
    parsed_dates = set()
    for date_format in formats:
        try:
            parsed_dates.add(datetime.strptime(date_text, date_format).date())
        except ValueError:
            continue

    if len(parsed_dates) == 1:
        return parsed_dates.pop().isoformat(), True, None
    if len(parsed_dates) > 1:
        return None, False, "Ambiguous date format"
    return None, False, error


def _invoice_columns(columns: Iterable[str]) -> Dict[str, Optional[str]]:
    """Map known invoice concepts to the matching source header, if present."""
    aliases = {
        "invoice_id": ["invoice_no", "invoice_number", "invoice_id", "invoice #", "invoice"],
        "order_date": ["order_date", "invoice_date", "order date", "invoice date", "date"],
        "customer_id": ["customer_id", "customer", "customer_name", "client", "client_name"],
        "product_id": ["product_id", "product", "product_name", "item", "item_name"],
        "quantity": ["quantity", "qty"],
        "price": ["unit_price", "price", "amount", "total", "line_total", "revenue"],
    }
    return {
        field: identify_column(columns, names)
        for field, names in aliases.items()
    }


def _is_invoice_data(columns: Iterable[str]) -> bool:
    """Recognize invoice-like headers without requiring customer contact fields.

    Requiring an invoice identifier plus at least one line-item/amount column
    helps keep ordinary customer lists on the customer-cleaning path.
    """
    invoice_columns = _invoice_columns(columns)
    has_product_or_amount = (
        invoice_columns["product_id"] is not None
        or invoice_columns["quantity"] is not None
        or invoice_columns["price"] is not None
    )
    return invoice_columns["invoice_id"] is not None and has_product_or_amount


def _write_spreadsheet(dataframe: pd.DataFrame, output_path: Path) -> None:
    """Write CSV, legacy XLS, or XLSX according to the output suffix."""
    extension = output_path.suffix.lower()
    if extension == ".csv":
        dataframe.to_csv(output_path, index=False)
    elif extension == ".xls":
        _write_xls(dataframe, output_path)
    else:
        dataframe.to_excel(output_path, index=False, engine="openpyxl")


def _clean_invoice_data(
    df: pd.DataFrame,
    path: Path,
    output_folder: str,
    dry_run: bool,
) -> Tuple[pd.DataFrame, CleaningReport]:
    """Normalize invoice rows and mark incomplete or invalid records for review.

    Unlike customer records, invoice rows do not require email or phone
    columns. The original source columns are retained, and this workflow
    currently routes questionable rows to review rather than rejecting them.
    """
    columns = _invoice_columns(df.columns)
    report = CleaningReport(original_file=path.name, total_rows=len(df))

    # Only recognized date and numeric columns are transformed; unrelated
    # source columns pass through unchanged apart from whitespace trimming.
    date_columns = {
        identify_column(df.columns, aliases)
        for aliases in (
            ["order_date", "invoice_date", "order date", "invoice date", "date"],
            ["ship_date", "shipping_date", "shipment_date", "ship date", "shipping date"],
        )
    }
    date_columns.discard(None)
    numeric_columns = {
        identify_column(df.columns, aliases)
        for aliases in (
            ["quantity", "qty"],
            ["unit_cost", "cost_per_unit", "unit cost"],
            ["unit_price", "price_per_unit", "unit price"],
            ["revenue", "amount", "line_total", "total"],
            ["cost"],
            ["margin", "profit"],
        )
    }
    numeric_columns.discard(None)
    output_columns = ["Original_Row", *df.columns, "Status"]
    cleaned_rows: List[Dict[str, Any]] = []
    review_rows: List[Dict[str, Any]] = []

    # A required concept is checked only when the input has a matching column.
    # For concepts with aliases, any populated alias satisfies the requirement.
    required_groups = [
        ("invoice number", [columns["invoice_id"]]),
        ("date", [columns["order_date"]]),
        ("customer", [
            identify_column(df.columns, [alias])
            for alias in ("customer_id", "customer", "customer_name", "client", "client_name")
        ]),
        ("product", [
            identify_column(df.columns, [alias])
            for alias in ("product_id", "product", "product_name", "item", "item_name")
        ]),
        ("quantity", [columns["quantity"]]),
        ("price or amount", [
            identify_column(df.columns, [alias])
            for alias in ("unit_price", "price", "amount", "total", "line_total", "revenue")
        ]),
    ]
    required_groups = [
        (label, [column for column in group if column is not None])
        for label, group in required_groups
        if any(column is not None for column in group)
    ]

    for source_row, (_, row) in enumerate(df.iterrows(), start=2):
        report.processed_rows += 1
        row_dict = row.to_dict()
        cleaned_row: Dict[str, Any] = {"Original_Row": source_row}
        notes: List[str] = []

        # Keep the original field names and values where normalization fails;
        # the notes explain why the row was routed to human review.
        for column, value in row_dict.items():
            if isinstance(value, str):
                value = re.sub(r"\s+", " ", value.strip())
            if column in date_columns and _has_value(value):
                normalized, valid, error = _normalize_invoice_date(value)
                if valid:
                    value = normalized
                else:
                    notes.append(f"{column}: {error}")
            elif column in numeric_columns and _has_value(value):
                normalized, valid, error = normalize_currency(value)
                if valid:
                    value = normalized
                else:
                    notes.append(f"{column}: {error}")
            cleaned_row[column] = value

        for label, group_columns in required_groups:
            if not any(_has_value(row_dict[column]) for column in group_columns):
                notes.append(f"Missing required {label}")

        if notes:
            cleaned_row["Status"] = "Review"
            cleaned_row["Cleaning_Notes"] = "; ".join(notes)
            review_rows.append(cleaned_row)
            report.review_count += 1
        else:
            cleaned_row["Status"] = "Clean"
            cleaned_rows.append(cleaned_row)
            report.cleaned_count += 1

    # A dry run performs all classification but deliberately writes no files.
    if not dry_run:
        out_path = Path(output_folder)
        out_path.mkdir(parents=True, exist_ok=True)
        base_ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        ts = base_ts
        suffix = 1
        while any(
            (out_path / f"{prefix}_{path.stem}_{ts}{extension}").exists()
            for prefix, extension in (
                ("cleaned", path.suffix.lower()),
                ("review_needed", ".csv"),
                ("rejected", ".csv"),
                ("report", ".txt"),
            )
        ):
            ts = f"{base_ts}_{suffix}"
            suffix += 1

        cleaned_df = pd.DataFrame(cleaned_rows, columns=output_columns)
        cleaned_path = out_path / f"cleaned_{path.stem}_{ts}{path.suffix.lower()}"
        _write_spreadsheet(cleaned_df, cleaned_path)

        review_columns = [*output_columns, "Cleaning_Notes"]
        review_df = pd.DataFrame(review_rows, columns=review_columns)
        review_df.to_csv(
            out_path / f"review_needed_{path.stem}_{ts}.csv",
            index=False,
        )
        pd.DataFrame(columns=["Original_Row", "Cleaning_Errors"]).to_csv(
            out_path / f"rejected_{path.stem}_{ts}.csv",
            index=False,
        )

        with open(
            out_path / f"report_{path.stem}_{ts}.txt",
            "w",
            encoding="utf-8",
        ) as report_file:
            report_file.write("INVOICE DATA CLEANING REPORT\n")
            report_file.write(f"File: {path.name} | Date: {datetime.now()}\n")
            report_file.write("=" * 40 + "\n\n")
            report_file.write("DATA QUALITY METRICS:\n")
            for metric, value in report.summary_stats().items():
                report_file.write(f"{metric}: {value:.2f}%\n")
            report_file.write(f"\nTotal Processed: {report.processed_rows}\n")
            report_file.write(f"Clean: {report.cleaned_count}\n")
            report_file.write(f"Needs Human Review: {report.review_count}\n")
            report_file.write("Rejected: 0\nDuplicates Removed: 0\n")

    return pd.DataFrame(cleaned_rows, columns=output_columns), report


def identify_column(cols: Iterable[str], target_keywords: List[str]) -> Optional[str]:
    """Find the first exact header alias after ignoring case and punctuation.

    Exact normalized matching avoids accidentally mapping a short alias to a
    different field whose name merely contains that alias.
    """
    normalized_targets = {
        re.sub(r"[^a-z0-9]", "", keyword.lower())
        for keyword in target_keywords
    }
    for col in cols:
        normalized_col = re.sub(r"[^a-z0-9]", "", str(col).lower())
        if normalized_col in normalized_targets:
            return col
    return None


def _write_xls(dataframe: pd.DataFrame, output_path: Path) -> None:
    """Write a legacy XLS workbook, converting cell values to xlwt-safe types."""
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("Cleaned Data")

    for column_index, column in enumerate(dataframe.columns):
        worksheet.write(0, column_index, str(column))

    # xlwt does not accept every pandas/numpy scalar, so convert values before
    # writing each cell and represent missing values as empty cells.
    for row_index, row in enumerate(dataframe.itertuples(index=False, name=None), start=1):
        for column_index, value in enumerate(row):
            if pd.isna(value):
                value = ""
            elif isinstance(value, (date, datetime)):
                value = value.isoformat()
            elif hasattr(value, "item"):
                value = value.item()
            if not isinstance(value, (str, bool, int, float)):
                value = str(value)
            worksheet.write(row_index, column_index, value)

    workbook.save(str(output_path))


def clean_data(file_path: str, output_folder: str, dry_run: bool = False) -> Tuple[pd.DataFrame, CleaningReport]:
    """Clean a CSV/XLS/XLSX file and optionally write result files.

    Invoice-shaped headers select the invoice workflow; other supported
    inputs use customer validation. Returns the clean rows and a report in
    both normal and dry-run modes.
    """
    # Load the source spreadsheet according to its file extension.
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    ext = path.suffix.lower()
    if ext == '.csv':
        # Preserve identifiers and literal strings such as "NA" and leading
        # zeroes; keep blank lines so source row numbers stay aligned.
        df = pd.read_csv(
            path,
            dtype=str,
            keep_default_na=False,
            skip_blank_lines=False,
        )
    elif ext in ['.xlsx', '.xls']:
        # Avoid treating legitimate text values such as "NA" as empty cells.
        df = pd.read_excel(path, keep_default_na=False)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    # Select a workflow from the headers, before applying customer-specific
    # requirements such as a name and a valid contact method.
    if _is_invoice_data(df.columns):
        return _clean_invoice_data(df, path, output_folder, dry_run)

    # Initialize counters and map input headers to the fields used below.
    report = CleaningReport(original_file=path.name, total_rows=len(df))

    # Identify likely source columns and ensure the required fields are present.
    col_map = {
        'name': identify_column(df.columns, ['name', 'customer', 'customer name', 'full name']),
        'email': identify_column(df.columns, ['email', 'e-mail', 'mail', 'email address']),
        'phone': identify_column(df.columns, ['phone', 'telephone', 'tel', 'mobile', 'phone number']),
        'dob': identify_column(df.columns, ['dob', 'date of birth', 'birth date', 'birthdate']),
        'revenue': identify_column(df.columns, ['revenue', 'amount', 'value', 'price']),
        'date': identify_column(df.columns, ['date', 'signup date', 'created at'])
    }

    # Name is required; contact details may use either email or phone, and
    # DOB, signup date, and revenue are optional.
    if col_map['name'] is None:
        raise ValueError("Required column not identified: name")
    if col_map['email'] is None and col_map['phone'] is None:
        raise ValueError("At least one contact column must be present: email or phone")

    # Reserve normalized output names so original extra columns cannot
    # overwrite generated status, row-number, or explanation fields. Unknown
    # source fields are retained under their own names where possible.
    standard_columns = [
        "Original_Row", "Name", "Email", "Phone", "DOB", "Date", "Revenue", "Status"
    ]
    original_value_columns = [
        "Original_Email",
        "Original_Phone",
        "Original_DOB",
        "Original_Date",
        "Original_Revenue",
    ]
    reserved_columns = set(
        standard_columns
        + ["Cleaning_Notes", "Cleaning_Errors"]
        + original_value_columns
    )
    passthrough_map: Dict[Any, str] = {}
    used_columns = set(reserved_columns)
    for col in df.columns:
        if col in col_map.values():
            continue
        output_col = str(col)
        while output_col in used_columns:
            output_col = f"Source_{output_col}"
        passthrough_map[col] = output_col
        used_columns.add(output_col)

    # Rejected output preserves every raw source column, so give conflicting
    # source headers a unique name before adding rejection metadata.
    rejected_reserved_columns = {"Original_Row", "Cleaning_Errors"}
    rejected_source_map: Dict[Any, str] = {}
    used_rejected_columns = set(rejected_reserved_columns)
    for col in df.columns:
        output_col = str(col)
        while output_col in used_rejected_columns:
            output_col = f"Source_{output_col}"
        rejected_source_map[col] = output_col
        used_rejected_columns.add(output_col)

    def rejected_output_row(row_dict: Dict[Any, Any], source_row: int, reason: str) -> Dict[str, Any]:
        """Keep rejected source values intact alongside row number and reason."""
        result = {
            output_col: row_dict[source_col]
            for source_col, output_col in rejected_source_map.items()
        }
        result["Original_Row"] = source_row
        result["Cleaning_Errors"] = reason
        return result

    cleaned_rows = []
    review_rows = []
    rejected_rows = []

    # Phone is the automatic duplicate key in this workflow. Register a phone
    # only after the row is otherwise accepted, so a rejected row cannot cause
    # a later usable row to be treated as its duplicate.
    phone_registry: Dict[str, int] = {} # normalized_phone -> source row index

    # Normalize each input row, recording validation problems before deciding
    # whether the row belongs in the clean, review, or rejected output.
    for source_row, (_, row) in enumerate(df.iterrows(), start=2):
        report.processed_rows += 1
        row_dict = row.to_dict()
        errors = []
        warnings = []

        # 1. Normalize Name
        raw_name = row_dict.get(col_map['name'])
        norm_name = normalize_name(raw_name)

        # 2. Normalize Email
        raw_email = row_dict.get(col_map['email'])
        norm_email, email_valid, email_err = normalize_email(raw_email)
        if _has_value(raw_email) and not email_valid:
            errors.append(f"Email Error: {email_err}")

        # 3. Normalize Phone
        raw_phone = row_dict.get(col_map['phone'])
        norm_phone, phone_valid, phone_err = normalize_phone(raw_phone)
        if _has_value(raw_phone) and not phone_valid:
            errors.append(f"Phone Error: {phone_err}")

        # 4. Normalize DOB
        raw_dob = row_dict.get(col_map['dob']) if col_map['dob'] else None
        norm_dob, dob_valid, dob_err = normalize_date(raw_dob)
        if _has_value(raw_dob) and not dob_valid:
            errors.append(f"DOB Error: {dob_err}")

        # 5. Normalize Signup Date
        raw_date = row_dict.get(col_map['date']) if col_map['date'] else None
        norm_date, date_valid, date_err = normalize_date(raw_date)
        if _has_value(raw_date) and not date_valid:
            errors.append(f"Date Error: {date_err}")

        # 6. Normalize Revenue
        raw_rev = row_dict.get(col_map['revenue']) if col_map['revenue'] else None
        norm_rev, rev_valid, rev_err = normalize_currency(raw_rev)
        if _has_value(raw_rev) and not rev_valid:
            errors.append(f"Revenue Error: {rev_err}")

        # Missing names and rows without any valid contact are hard failures.
        # Other field errors can be reviewed if the row still has a name and
        # at least one usable contact method.
        has_valid_email = email_valid and norm_email is not None
        has_valid_phone = phone_valid and norm_phone is not None
        if not norm_name:
            errors.append("Name Error: Missing or invalid name")
        if not (has_valid_email or has_valid_phone):
            errors.append("Contact Error: No valid email or phone number")

        # Reject rows meeting the hard-rejection condition; otherwise retain
        # non-critical validation errors as review warnings.
        if errors:
            if not norm_name or not (has_valid_email or has_valid_phone):
                report.rejected_count += 1
                report.rejected_rows.append({"row": source_row, "reason": "; ".join(errors)})
                rejected_rows.append(rejected_output_row(row_dict, source_row, "; ".join(errors)))
                continue
            # Soft errors can be retained for human review if a valid contact exists.
            warnings.append("; ".join(errors))

        if not norm_email and not norm_phone:
            warnings.append("Missing contact info (Email or Phone)")

        # Keep invalid raw values in the review output so a person can correct
        # them without having to consult the original file.
        original_values = {}
        for field_name, raw_value, is_valid in (
            ("Email", raw_email, email_valid),
            ("Phone", raw_phone, phone_valid),
            ("DOB", raw_dob, dob_valid),
            ("Date", raw_date, date_valid),
            ("Revenue", raw_rev, rev_valid),
        ):
            if _has_value(raw_value) and not is_valid:
                original_values[f"Original_{field_name}"] = raw_value

        # Repeated normalized phone numbers are removed from the clean/review
        # sets and retained in the rejected file with an explanation.
        is_duplicate = False
        dup_reason = ""

        if norm_phone and norm_phone in phone_registry:
            is_duplicate = True
            dup_reason = f"Duplicate phone {norm_phone}"

        if is_duplicate:
            report.duplicates_removed += 1
            report.rejected_rows.append({"row": source_row, "reason": dup_reason})
            rejected_rows.append(rejected_output_row(row_dict, source_row, dup_reason))
            continue
        if norm_phone:
            phone_registry[norm_phone] = source_row

        # Build the normalized output record and retain all unrecognized columns.
        cleaned_row = {
            'Original_Row': source_row,
            'Name': norm_name,
            'Email': norm_email,
            'Phone': norm_phone,
            'DOB': norm_dob,
            'Date': norm_date,
            'Revenue': norm_rev,
            'Status': 'Clean' if not warnings else 'Review'
        }
        for source_col, output_col in passthrough_map.items():
            cleaned_row[output_col] = row_dict[source_col]

        if warnings:
            cleaned_row.update(original_values)
            cleaned_row['Cleaning_Notes'] = "; ".join(warnings)
            report.review_count += 1
            review_rows.append(cleaned_row)
        else:
            report.cleaned_count += 1
            cleaned_rows.append(cleaned_row)

    # Similar names sharing a DOB are only potential matches. They are reported
    # for human review and are never automatically removed by this scan.
    all_kept = cleaned_rows + review_rows

    # Group by DOB first so comparisons are limited to people with the same
    # normalized date of birth instead of comparing every pair in the file.
    dob_groups: Dict[Optional[str], List[Dict[str, Any]]] = {}
    for row in all_kept:
        dob = row.get('DOB')
        if dob:
            dob_groups.setdefault(dob, []).append(row)

    # Both caps bound runtime and report size for unusually large DOB groups.
    # The report flag makes it explicit when the scan may be incomplete.
    candidate_limit = 10000
    comparison_limit = 100000
    comparison_count = 0
    candidate_scan_limited = False
    for dob, members in dob_groups.items():
        # Within each DOB group, perform a bounded fuzzy comparison.
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                if comparison_count >= comparison_limit:
                    candidate_scan_limited = True
                    break
                comparison_count += 1
                name1 = get_comparison_name(members[i]['Name'])
                name2 = get_comparison_name(members[j]['Name'])
                if name1 and name2:
                    score = fuzz.token_sort_ratio(name1, name2)
                    if score >= 85:
                        report.fuzzy_matches.append({
                            "row_a": members[i]['Original_Row'],
                            "row_b": members[j]['Original_Row'],
                            "score": score,
                            "matching_field": "DOB",
                            "names": f"{name1} vs {name2}"
                        })
                        if len(report.fuzzy_matches) >= candidate_limit:
                            candidate_scan_limited = True
                            break
            if candidate_scan_limited:
                break
        if candidate_scan_limited:
            break

    report.potential_matches_limited = candidate_scan_limited

    # Generate a unique timestamp for this run so prior exports are not
    # overwritten, then write clean, review, rejected, and summary outputs.
    if not dry_run:
        out_path = Path(output_folder)
        out_path.mkdir(parents=True, exist_ok=True)
        base_ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        ts = base_ts
        suffix = 1
        while any(
            (out_path / f"{prefix}_{path.stem}_{ts}{extension}").exists()
            for prefix, extension in (
                ("cleaned", path.suffix.lower()),
                ("review_needed", ".csv"),
                ("rejected", ".csv"),
                ("report", ".txt"),
            )
        ):
            ts = f"{base_ts}_{suffix}"
            suffix += 1

        passthrough_columns = list(passthrough_map.values())

        # Always write the clean output, including a header-only file when no
        # rows passed validation, so each run has a predictable output set.
        cleaned_df = pd.DataFrame(cleaned_rows)
        cols_to_keep = standard_columns + passthrough_columns
        if cleaned_df.empty:
            cleaned_df = pd.DataFrame(columns=cols_to_keep)
        else:
            cleaned_df = cleaned_df.reindex(columns=cols_to_keep)
        cleaned_path = out_path / f"cleaned_{path.stem}_{ts}{path.suffix.lower()}"
        _write_spreadsheet(cleaned_df, cleaned_path)

        # Review exports include validation notes and original invalid values.
        review_df = pd.DataFrame(review_rows)
        review_cols = (
            standard_columns
            + ["Cleaning_Notes"]
            + original_value_columns
            + passthrough_columns
        )
        if review_df.empty:
            review_df = pd.DataFrame(columns=review_cols)
        else:
            review_df = review_df.reindex(columns=review_cols)
        review_df.to_csv(out_path / f"review_needed_{path.stem}_{ts}.csv", index=False)

        # Rejected exports preserve the raw source columns for correction or
        # audit, while the reason explains why each row was excluded.
        rejected_df = pd.DataFrame(rejected_rows)
        if rejected_df.empty:
            rejected_columns = (
                list(rejected_source_map.values())
            )
            rejected_columns += ["Original_Row", "Cleaning_Errors"]
            rejected_df = pd.DataFrame(columns=rejected_columns)
        rejected_df.to_csv(out_path / f"rejected_{path.stem}_{ts}.csv", index=False)

        # The text report summarizes outcomes and lists the human-review
        # candidates and rejection reasons found during this run.
        with open(out_path / f"report_{path.stem}_{ts}.txt", 'w', encoding='utf-8') as f:
            stats = report.summary_stats()
            f.write(f"PROFESSIONAL DATA CLEANING REPORT\n")
            f.write(f"File: {path.name} | Date: {datetime.now()}\n")
            f.write("="*40 + "\n\n")
            f.write(f"DATA QUALITY METRICS:\n")
            for k, v in stats.items():
                f.write(f"{k}: {v:.2f}%\n")
            f.write(f"\nTotal Processed: {report.processed_rows}\n")
            f.write(f"Perfectly Clean: {report.cleaned_count}\n")
            f.write(f"Needs Human Review: {report.review_count}\n")
            f.write(f"Rejected: {report.rejected_count}\n")
            f.write(f"Duplicates Removed: {report.duplicates_removed}\n")

            f.write("\n" + "="*40 + "\n")
            f.write("POTENTIAL DUPLICATES (Shared DOB; requires human review):\n")
            if not report.fuzzy_matches:
                f.write("No potential matches found.\n")
            for match in report.fuzzy_matches:
                f.write(f"Rows {match['row_a']} & {match['row_b']} | Score: {match['score']} | {match['names']}\n")
            if report.potential_matches_limited:
                f.write("Candidate scan was limited; additional matches may exist.\n")

            f.write("\n" + "="*40 + "\n")
            f.write("REJECTION LOG:\n")
            for rej in report.rejected_rows:
                f.write(f"Row {rej['row']}: {rej['reason']}\n")

    return pd.DataFrame(cleaned_rows), report


if __name__ == "__main__":
    # This block runs only when invoked as a script; importing Cleaner exposes
    # the cleaning helpers without starting a command-line workflow.
    # Parse command-line options for the input file, output folder, and preview.
    parser = argparse.ArgumentParser(description="Spreadsheet Data Cleaning Pipeline")
    parser.add_argument("input", help="Path to input CSV/Excel")
    parser.add_argument("--output", default="cleaned_output", help="Output folder")
    parser.add_argument("--dry-run", action="store_true", help="Preview changes")

    args = parser.parse_args()
    # Run the workflow and display a concise summary; report failures to both
    # the configured logger and the command-line user.
    try:
        df, report = clean_data(args.input, args.output, args.dry_run)
        print(f"\n--- Workflow Summary ---")
        print(f"Clean: {report.cleaned_count} | Review: {report.review_count} | Rejected: {report.rejected_count}")
        print(f"Duplicates: {report.duplicates_removed}")
        if not args.dry_run:
            print(f"Outputs saved to: {args.output}")
    except Exception as e:
        logger.error(f"Error: {e}")
        print(f"Error: {e}")
        sys.exit(1)
