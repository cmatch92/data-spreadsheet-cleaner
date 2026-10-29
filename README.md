# 🧹 Spreadsheet Data Cleaner

<p align="center">
  <strong>A Python command-line tool for cleaning customer and invoice spreadsheets</strong><br>
  Normalize common fields, review questionable rows, and keep a record of each run.
</p>

---

## 📖 About

Spreadsheet Data Cleaner reads CSV and Excel files, recognizes common customer
or invoice columns, and sorts rows into clean and review outputs. Customer rows
that cannot be reliably used are also recorded in a rejected output. The tool
preserves unrecognized columns so additional source data is not discarded.

This is a local file-processing tool. It does **not** connect to or modify a
live database.

## ✨ Features

- 📄 **CSV and Excel support:** read `.csv`, `.xls`, and `.xlsx` files, and
  write the clean spreadsheet in the same format as the input.
- 🧹 **Customer-field normalization:** standardize names, validate email
  syntax, format valid phone numbers, normalize supported dates, and parse
  currency values.
- 🧾 **Invoice workflow:** detect invoice-style headers and clean date,
  numeric, and text values without requiring customer contact fields.
- 🔎 **Review instead of guessing:** send questionable rows to a separate
  review file with notes and, where applicable, the original invalid values.
- ♻️ **Duplicate checks:** repeated normalized customer phone numbers are
  excluded from the clean/review files and listed in the rejected file.
- 👀 **Potential-match report:** similar names sharing a date of birth are
  reported for human review; they are not automatically removed.
- 📊 **Run reports:** save a text report with row counts, quality percentages,
  review candidates, and rejection reasons.
- 🧪 **Automated tests:** cover normalization, customer and invoice workflows,
  spreadsheet formats, and command-line error handling.

## 🚀 Get started

### Requirements

- Python 3.11 or newer recommended
- `pandas`, `thefuzz`, `email-validator`, `phonenumbers`, `openpyxl`,
  `xlrd`, and `xlwt`
- `pytest` to run the test suite

### Install

Clone or download this repository, open a terminal in the project directory,
and install the dependencies:

```bash
python -m pip install "pandas>=2.2,<4" "thefuzz>=0.22,<1" "email-validator>=2.2,<3" "phonenumbers>=8.13,<10" "openpyxl>=3.1,<4" "xlrd>=2,<3" "xlwt>=1.3,<2" "pytest>=8,<9"
```

### Run the cleaner

Clean a file and write results to the default `cleaned_output` folder:

```bash
python Cleaner.py path/to/customers.csv
```

Choose a different output folder:

```bash
python Cleaner.py path/to/invoices.xlsx --output path/to/results
```

Preview the classification without writing output files:

```bash
python Cleaner.py path/to/customers.csv --dry-run
```

Use `python Cleaner.py --help` to see the available command-line options.

## 🗂️ Input data

### Customer files

The customer workflow needs a name column and at least one email or phone
column. Headers are matched against known aliases without regard to case,
spaces, or punctuation.

| Field | Example recognized headers | Requirement |
| --- | --- | --- |
| Name | `name`, `customer name`, `full name` | Required |
| Email | `email`, `e-mail`, `email address` | Email or phone column required |
| Phone | `phone`, `telephone`, `mobile`, `phone number` | Email or phone column required |
| Date of birth | `dob`, `date of birth`, `birthdate` | Optional |
| Signup date | `date`, `signup date`, `created at` | Optional |
| Revenue | `revenue`, `amount`, `value`, `price` | Optional |

Other source columns are retained in the output. Customer rows missing a name
or every valid contact method are rejected. Invalid optional values can be
preserved in the review output for correction.

### Invoice files

The invoice workflow is selected when the headers include an invoice identifier
and at least one product, quantity, or amount field. Common headers such as
`invoice_no`, `order_date`, `customer_id`, `product_id`, `quantity`, and
`unit_price` are recognized. Rows with missing values in recognized required
fields or invalid supported dates/numbers are marked for review.

Date formats are interpreted differently by workflow: ambiguous slash-formatted
customer dates are flagged rather than guessed, while invoice slash-formatted
dates use month-first interpretation.

## 📦 Output files

Each run that writes results creates uniquely named files in the selected
output folder:

| File | Contents |
| --- | --- |
| `cleaned_<input>_<timestamp>` | Rows classified as clean, in the input spreadsheet format |
| `review_needed_<input>_<timestamp>.csv` | Rows requiring human review and notes |
| `rejected_<input>_<timestamp>.csv` | Rejected customer rows with their source values and reasons |
| `report_<input>_<timestamp>.txt` | Counts, quality metrics, potential matches, and rejection log |

Invoice rows are routed to clean or review; the invoice workflow does not
currently reject rows.

## 🧪 Run the tests

After installing the dependencies above, run:

```bash
python -m pytest -q
```

## 🔐 Data handling

The input and generated files may contain personal or confidential business
information. Process them only in an environment authorized for that data,
restrict access to the output folder, and do not commit real customer or
invoice files to a public repository. Review all suggested duplicate matches
and rejected rows before making changes to a source system.

## 📁 Project structure

```text
Cleaner.py             Cleaning workflows and command-line interface
test_cleaner.py        Automated tests
.gitignore             Excludes local data and generated files
README.md              Project documentation
```

## 🛠️ Built with

- **Python** — command-line application and cleaning logic
- **pandas** — CSV and Excel data handling
- **email-validator** and **phonenumbers** — contact-field validation
- **thefuzz** — potential-match comparisons
- **pytest** — automated tests
