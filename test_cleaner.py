import pandas as pd
import subprocess
import sys
from pathlib import Path

from Cleaner import (
    clean_data,
    normalize_currency,
    normalize_date,
    normalize_email,
    normalize_name,
    normalize_phone,
)


def test_normalize_name():
    assert normalize_name("  john smith  ") == "John Smith"
    assert normalize_name("JOHN SMITH") == "John Smith"
    assert normalize_name("jAnE   dOe") == "Jane Doe"
    assert normalize_name(None) is None
    assert normalize_name("") is None


def test_normalize_email():
    email, valid, error = normalize_email("  JOHN@GMAIL.com  ")
    assert email == "john@gmail.com"
    assert valid is True
    assert error is None

    email, valid, error = normalize_email("invalid-email")
    assert email is None
    assert valid is False
    assert error


def test_normalize_phone():
    phone, valid, error = normalize_phone("(650) 253-0000")
    assert phone == "+16502530000"
    assert valid is True
    assert error is None

    phone, valid, error = normalize_phone("555-1234")
    assert phone is None
    assert valid is False
    assert error

    assert normalize_phone(None) == (None, True, None)


def test_normalize_date_rejects_ambiguous_values():
    date, valid, error = normalize_date("2023-01-15")
    assert date == "2023-01-15" and valid is True and error is None

    date, valid, error = normalize_date("01/15/2023")
    assert date == "2023-01-15" and valid is True and error is None

    date, valid, error = normalize_date("04/05/2023")
    assert date is None and valid is False
    assert error is not None
    assert "Ambiguous" in error

    date, valid, error = normalize_date("not a date")
    assert date is None and valid is False
    assert error

    date, valid, error = normalize_date(pd.Timestamp("2023-01-15"))
    assert date == "2023-01-15" and valid is True and error is None

    date, valid, error = normalize_date(45292)
    assert date == "2024-01-01" and valid is True and error is None


def test_normalize_currency_validates_grouping_and_finite_values():
    value, valid, error = normalize_currency("$1,500.00")
    assert value == 1500.0 and valid is True and error is None

    value, valid, error = normalize_currency("$ 1,500.00")
    assert value == 1500.0 and valid is True and error is None

    value, valid, error = normalize_currency("(25.00)")
    assert value == -25.0 and valid is True and error is None

    for malformed in ("1,2,3", "$1$", "NaN", "Infinity", "invalid"):
        value, valid, error = normalize_currency(malformed)
        assert value is None and valid is False
        assert error


def test_cleaning_workflow_triages_rows_and_reports_dob_candidates(tmp_path):
    csv_path = tmp_path / "test_messy.csv"
    csv_path.write_text(
        "Name,Email,Phone,DOB,Date,Revenue\n"
        "john smith,john@gmail.com,6502530000,1990-01-01,2023-01-01,$100\n"
        "JOHN SMITH,JOHN@GMAIL.COM,+1 6502530000,1990-01-01,2023-01-01,100\n"
        "john smyth,jane@gmail.com,6502530001,1990-01-01,2023-01-01,$200\n"
        "bad email,bad@,6502530002,1990-01-01,2023-01-01,100\n"
        "bob builder,bob@bob.com,6502530003,not-a-date,2023-01-01,100\n",
        encoding="utf-8",
    )

    output_folder = tmp_path / "output"
    cleaned, report = clean_data(str(csv_path), str(output_folder))

    assert report.processed_rows == 5
    assert report.cleaned_count == 2
    assert report.review_count == 2
    assert report.rejected_count == 0
    assert report.duplicates_removed == 1
    assert len(report.rejected_rows) == 1
    assert len(cleaned) == 2
    assert cleaned["Name"].tolist() == ["John Smith", "John Smyth"]
    assert any(
        item["row_a"] == 2 and item["row_b"] == 4
        and item["matching_field"] == "DOB"
        for item in report.fuzzy_matches
    )

    cleaned_file = next(output_folder.glob("cleaned_*.csv"))
    review_file = next(output_folder.glob("review_needed_*.csv"))
    rejected_file = next(output_folder.glob("rejected_*.csv"))
    assert len(pd.read_csv(cleaned_file)) == 2
    review_data = pd.read_csv(review_file)
    assert len(review_data) == 2
    assert review_data["Cleaning_Notes"].notna().all()
    rejected_data = pd.read_csv(rejected_file)
    assert len(rejected_data) == 1
    assert rejected_data.iloc[0]["Original_Row"] == 3
    assert len(list(output_folder.glob("report_*.txt"))) == 1


def test_shared_dob_with_unrelated_names_is_not_potential_duplicate(tmp_path):
    csv_path = tmp_path / "same_dob.csv"
    csv_path.write_text(
        "Name,Email,Phone,DOB\n"
        "Alice Smith,alice@example.com,6502530000,1990-01-01\n"
        "Robert Johnson,robert@example.com,6502530001,1990-01-01\n",
        encoding="utf-8",
    )

    _, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.fuzzy_matches == []


def test_optional_dob_and_date_columns_and_name_alone_not_duplicate(tmp_path):
    csv_path = tmp_path / "without_dob.csv"
    csv_path.write_text(
        "Customer Name,E-mail,Telephone\n"
        "Alex Smith,alex@example.com,6502530000\n"
        "Alex Smith,alex2@example.com,6502530001\n",
        encoding="utf-8",
    )

    cleaned, report = clean_data(
        str(csv_path),
        str(tmp_path / "output"),
        dry_run=True,
    )

    assert len(cleaned) == 2
    assert report.cleaned_count == 2
    assert report.duplicates_removed == 0
    assert report.fuzzy_matches == []


def test_email_only_customer_file_is_supported(tmp_path):
    csv_path = tmp_path / "email_only.csv"
    csv_path.write_text(
        "Name,Email\n"
        "Alex Smith,alex@example.com\n",
        encoding="utf-8",
    )

    _, report = clean_data(
        str(csv_path),
        str(tmp_path / "output"),
        dry_run=True,
    )

    assert report.processed_rows == 1
    assert report.cleaned_count == 1
    assert report.review_count == 0
    assert report.rejected_count == 0


def test_rows_without_any_valid_contact_are_rejected(tmp_path):
    csv_path = tmp_path / "no_contact.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Alex Smith,,\n",
        encoding="utf-8",
    )

    _, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.rejected_count == 1
    assert report.review_count == 0
    assert "Contact Error" in report.rejected_rows[0]["reason"]


def test_invalid_email_is_not_written_as_a_normalized_value(tmp_path):
    csv_path = tmp_path / "invalid_email_review.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Alex Smith,not-an-email,6502530000\n",
        encoding="utf-8",
    )

    clean_data(str(csv_path), str(tmp_path / "output"))

    review_file = next((tmp_path / "output").glob("review_needed_*.csv"))
    review = pd.read_csv(review_file)
    assert pd.isna(review.iloc[0]["Email"])
    assert review.iloc[0]["Original_Email"] == "not-an-email"


def test_customers_may_share_email_without_being_removed_as_duplicates(tmp_path):
    csv_path = tmp_path / "shared_email.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Alex Smith,family@example.com,6502530000\n"
        "Jordan Smith,family@example.com,6502530001\n",
        encoding="utf-8",
    )

    cleaned, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert len(cleaned) == 2
    assert report.duplicates_removed == 0


def test_literal_na_text_is_not_loaded_as_a_missing_value(tmp_path):
    csv_path = tmp_path / "literal_na.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "NA,na@example.com,6502530000\n",
        encoding="utf-8",
    )

    cleaned, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.cleaned_count == 1
    assert cleaned.iloc[0]["Name"] == "Na"


def test_blank_csv_lines_do_not_shift_reported_source_rows(tmp_path):
    csv_path = tmp_path / "blank_line.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Alex Smith,alex@example.com,6502530000\n"
        "\n"
        ",invalid-email,\n",
        encoding="utf-8",
    )

    _, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.processed_rows == 3
    assert report.rejected_rows[-1]["row"] == 4


def test_review_output_preserves_invalid_original_values(tmp_path):
    csv_path = tmp_path / "invalid_values.csv"
    csv_path.write_text(
        "Name,Email,Phone,DOB,Signup Date,Revenue\n"
        'Alex Smith,alex@example.com,555-1234,not-a-date,not-a-date,"1,2"\n',
        encoding="utf-8",
    )

    clean_data(str(csv_path), str(tmp_path / "output"))

    review_file = next((tmp_path / "output").glob("review_needed_*.csv"))
    review = pd.read_csv(review_file)
    assert review.iloc[0]["Original_Phone"] == "555-1234"
    assert review.iloc[0]["Original_DOB"] == "not-a-date"
    assert review.iloc[0]["Original_Date"] == "not-a-date"
    assert review.iloc[0]["Original_Revenue"] == "1,2"


def test_invalid_row_is_rejected_with_source_row_number(tmp_path):
    csv_path = tmp_path / "invalid.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Missing contact,invalid-email,\n",
        encoding="utf-8",
    )

    _, report = clean_data(str(csv_path), str(tmp_path / "output"))

    assert report.rejected_count == 1
    rejected_file = next((tmp_path / "output").glob("rejected_*.csv"))
    rejected = pd.read_csv(rejected_file)
    assert rejected.iloc[0]["Original_Row"] == 2
    assert "Email Error" in rejected.iloc[0]["Cleaning_Errors"]


def test_missing_name_is_rejected_even_when_contact_is_valid(tmp_path):
    csv_path = tmp_path / "missing_name.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        ",alex@example.com,6502530000\n",
        encoding="utf-8",
    )

    _, report = clean_data(str(csv_path), str(tmp_path / "output"))

    assert report.rejected_count == 1
    rejected_file = next((tmp_path / "output").glob("rejected_*.csv"))
    rejected = pd.read_csv(rejected_file)
    assert rejected.iloc[0]["Original_Row"] == 2
    assert "Name Error" in rejected.iloc[0]["Cleaning_Errors"]


def test_output_names_do_not_overwrite_prior_run(tmp_path):
    csv_path = tmp_path / "input.csv"
    csv_path.write_text(
        "Name,Email,Phone\n"
        "Alex Smith,alex@example.com,6502530000\n",
        encoding="utf-8",
    )
    output_folder = tmp_path / "output"

    clean_data(str(csv_path), str(output_folder))
    clean_data(str(csv_path), str(output_folder))

    assert len(list(output_folder.glob("report_*.txt"))) == 2


def test_source_columns_cannot_overwrite_generated_output_fields(tmp_path):
    csv_path = tmp_path / "reserved_columns.csv"
    csv_path.write_text(
        "Name,Email,Phone,Status,Cleaning_Notes\n"
        "Alex Smith,alex@example.com,6502530000,Source Status,Source note\n",
        encoding="utf-8",
    )

    clean_data(str(csv_path), str(tmp_path / "output"))

    cleaned = pd.read_csv(next((tmp_path / "output").glob("cleaned_*.csv")))
    assert cleaned.iloc[0]["Status"] == "Clean"
    assert cleaned.iloc[0]["Source_Status"] == "Source Status"
    assert cleaned.iloc[0]["Source_Cleaning_Notes"] == "Source note"


def test_source_columns_cannot_overwrite_rejection_metadata(tmp_path):
    csv_path = tmp_path / "rejected_reserved_columns.csv"
    csv_path.write_text(
        "Name,Email,Phone,Original_Row,Cleaning_Errors\n"
        ",alex@example.com,6502530000,source-row,source-error\n",
        encoding="utf-8",
    )

    clean_data(str(csv_path), str(tmp_path / "output"))

    rejected_file = next((tmp_path / "output").glob("rejected_*.csv"))
    rejected = pd.read_csv(rejected_file)
    assert rejected.iloc[0]["Original_Row"] == 2
    assert rejected.iloc[0]["Cleaning_Errors"] == "Name Error: Missing or invalid name"
    assert rejected.iloc[0]["Source_Original_Row"] == "source-row"
    assert rejected.iloc[0]["Source_Cleaning_Errors"] == "source-error"


def test_csv_identifiers_with_leading_zero_are_preserved(tmp_path):
    csv_path = tmp_path / "leading_zero.csv"
    csv_path.write_text(
        "Name,Email,Phone,Customer Code\n"
        "Alex Smith,alex@example.com,6502530000,000123\n",
        encoding="utf-8",
    )

    cleaned, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.processed_rows == 1
    assert cleaned.iloc[0]["Customer Code"] == "000123"


def test_xlsx_input_is_supported(tmp_path):
    xlsx_path = tmp_path / "input.xlsx"
    pd.DataFrame(
        {
            "Name": ["Alex Smith"],
            "Email": ["alex@example.com"],
            "Phone": [6502530000.0],
            "Signup Date": [45292],
        }
    ).to_excel(xlsx_path, index=False)

    cleaned, report = clean_data(str(xlsx_path), str(tmp_path / "output"), dry_run=True)

    assert report.processed_rows == 1
    assert cleaned.iloc[0]["Phone"] == "+16502530000"
    assert cleaned.iloc[0]["Date"] == "2024-01-01"


def test_xlsx_input_writes_xlsx_output(tmp_path):
    xlsx_path = tmp_path / "input.xlsx"
    pd.DataFrame(
        {
            "Name": ["Alex Smith"],
            "Email": ["alex@example.com"],
            "Phone": [6502530000],
        }
    ).to_excel(xlsx_path, index=False)
    output_folder = tmp_path / "output"

    clean_data(str(xlsx_path), str(output_folder))

    cleaned_path = next(output_folder.glob("cleaned_*.xlsx"))
    cleaned = pd.read_excel(cleaned_path, dtype=str)
    assert cleaned.iloc[0]["Name"] == "Alex Smith"
    assert cleaned.iloc[0]["Email"] == "alex@example.com"


def test_invoice_workbook_is_cleaned_without_contact_columns(tmp_path):
    invoice_path = tmp_path / "invoices.xlsx"
    pd.DataFrame(
        [
            {
                "invoice_no": 101,
                "order_date": "28/10/2026",
                "customer_id": 7,
                "customer": "  ACME   Supplies  ",
                "product_id": 25,
                "product": "  Paper  ",
                "quantity": "2",
                "unit_price": "",
                "revenue": "$20.00",
                "ship_date": "10-29-2026",
            },
            {
                "invoice_no": "",
                "order_date": "2026-10-28",
                "customer_id": 8,
                "product_id": 26,
                "quantity": 1,
                "unit_price": 5,
                "revenue": 5,
            },
            {
                "invoice_no": 103,
                "order_date": "not a date",
                "customer_id": 9,
                "product_id": 27,
                "quantity": "not-a-number",
                "unit_price": 5,
                "revenue": 5,
            },
        ]
    ).to_excel(invoice_path, index=False)
    output_folder = tmp_path / "output"

    cleaned, report = clean_data(str(invoice_path), str(output_folder))

    assert report.processed_rows == 3
    assert report.cleaned_count == 1
    assert report.review_count == 2
    assert report.rejected_count == 0
    assert len(cleaned) == 1
    assert cleaned.iloc[0]["order_date"] == "2026-10-28"
    assert cleaned.iloc[0]["ship_date"] == "2026-10-29"
    assert cleaned.iloc[0]["customer"] == "ACME Supplies"
    assert cleaned.iloc[0]["revenue"] == 20.0
    assert cleaned.iloc[0]["Status"] == "Clean"

    cleaned_file = next(output_folder.glob("cleaned_*.xlsx"))
    review_file = next(output_folder.glob("review_needed_*.csv"))
    assert len(pd.read_excel(cleaned_file)) == 1
    review = pd.read_csv(review_file)
    assert len(review) == 2
    assert review["Cleaning_Notes"].str.contains("Missing required invoice number").any()
    assert review["Cleaning_Notes"].str.contains("order_date").any()
    assert review["Cleaning_Notes"].str.contains("quantity").any()


def test_invoice_ambiguous_dates_use_month_first_format():
    from Cleaner import _normalize_invoice_date

    assert _normalize_invoice_date("10/09/2026") == ("2026-10-09", True, None)
    assert _normalize_invoice_date("10-09-2026") == ("2026-10-09", True, None)


def test_xls_input_writes_xls_output(tmp_path):
    import xlwt

    xls_path = tmp_path / "input.xls"
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("Customers")
    for column, value in enumerate(("Name", "Email", "Phone")):
        worksheet.write(0, column, value)
    for column, value in enumerate(("Alex Smith", "alex@example.com", "6502530000")):
        worksheet.write(1, column, value)
    workbook.save(str(xls_path))
    output_folder = tmp_path / "output"

    clean_data(str(xls_path), str(output_folder))

    cleaned_path = next(output_folder.glob("cleaned_*.xls"))
    cleaned = pd.read_excel(cleaned_path, dtype=str)
    assert cleaned.iloc[0]["Name"] == "Alex Smith"
    assert cleaned.iloc[0]["Email"] == "alex@example.com"
    assert cleaned.iloc[0]["Phone"] == "+16502530000"


def test_fuzzy_duplicate_scan_is_bounded_and_reported(tmp_path):
    csv_path = tmp_path / "large_dob_group.csv"
    pd.DataFrame(
        {
            "Name": ["Alex Smith"] * 150,
            "Email": [f"customer{index}@example.com" for index in range(150)],
            "Phone": [f"650253{index:04d}" for index in range(150)],
            "DOB": ["1990-01-01"] * 150,
        }
    ).to_csv(csv_path, index=False)

    _, report = clean_data(str(csv_path), str(tmp_path / "output"), dry_run=True)

    assert report.potential_matches_limited is True
    assert len(report.fuzzy_matches) == 10000


def test_cli_returns_nonzero_when_input_is_missing():
    cleaner_path = Path(__file__).with_name("Cleaner.py")
    result = subprocess.run(
        [sys.executable, str(cleaner_path), "missing-input.csv", "--dry-run"],
        capture_output=True,
        text=True,
        check=False,
    )

    assert result.returncode != 0
    assert "File not found" in result.stdout
