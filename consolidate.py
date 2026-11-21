"""Consolidate messy sales exports into one clean Excel report.

Usage:
    python consolidate.py --input sample_data --output output/sales_report.xlsx

Every .csv and .xlsx file in the input folder is treated as one store's export.
The script works out each file's delimiter, header position, column names,
date format and number format on its own, so new stores can be added without
changing the code.
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Column names seen in the wild, mapped to the field they mean.
ALIASES = {
    "date": "date", "order date": "date",
    "product": "product", "item": "product", "product name": "product",
    "qty": "quantity", "quantity": "quantity", "units": "quantity",
    "unit price": "unit_price", "price": "unit_price", "price usd": "unit_price",
    "customer": "customer", "client": "customer", "customer name": "customer",
}
FIELDS = ["date", "product", "quantity", "unit_price", "customer"]
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%b %d, %Y", "%Y-%m-%d %H:%M:%S"]
MONEY_FORMAT = "#,##0.00"


def normalize_header(text):
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def load_raw(path):
    """Read a file as a grid of trimmed strings, with no assumptions about layout."""
    if path.suffix.lower() == ".xlsx":
        raw = pd.read_excel(path, header=None, dtype=object)
    else:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        first_line = next((line for line in lines if line.strip()), "")
        delimiter = max(",;\t|", key=first_line.count)
        raw = pd.read_csv(
            path, header=None, dtype=str, sep=delimiter,
            skip_blank_lines=False, encoding="utf-8-sig",
        )
    return raw.map(lambda v: "" if pd.isna(v) else str(v).strip())


def locate_header(raw):
    """Find the first row that contains every expected column, in any order."""
    for index, row in raw.iterrows():
        mapping = {}
        for column, cell in row.items():
            field = ALIASES.get(normalize_header(cell))
            if field:
                mapping[column] = field
        if set(mapping.values()) == set(FIELDS):
            return index, mapping
    raise ValueError("no header row with date, product, quantity, price and customer")


def parse_dates(values):
    """Try each known format and keep the one that understands the most rows."""
    best = None
    for date_format in DATE_FORMATS:
        parsed = pd.to_datetime(values, format=date_format, errors="coerce")
        if best is None or parsed.notna().sum() > best.notna().sum():
            best = parsed
    return best


def parse_number(text):
    """Turn '$1,249.00', '1.249,00' or '24,90' into a float. Returns None if not a number."""
    cleaned = re.sub(r"[^\d,.\-]", "", text)
    if not re.search(r"\d", cleaned):
        return None
    if "," in cleaned and "." in cleaned:
        # Whichever separator comes last is the decimal one.
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        if re.fullmatch(r"-?\d+,\d{1,2}", cleaned):
            cleaned = cleaned.replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def rejection_reason(row):
    if not row["product"]:
        return "missing product"
    if pd.isna(row["date"]):
        return "invalid date"
    quantity = row["quantity"]
    if pd.isna(quantity) or quantity <= 0 or quantity != int(quantity):
        return "invalid quantity"
    if pd.isna(row["unit_price"]) or row["unit_price"] <= 0:
        return "invalid price"
    return ""


def load_store(path):
    """Return (valid rows, rejected rows) for one store export."""
    raw = load_raw(path)
    header_index, mapping = locate_header(raw)

    data = raw.loc[header_index + 1:, list(mapping)].rename(columns=mapping)[FIELDS]
    data = data[(data != "").any(axis=1)].copy()
    original = data.copy()

    data["date"] = parse_dates(data["date"])
    data["quantity"] = data["quantity"].map(parse_number).astype(float)
    data["unit_price"] = data["unit_price"].map(parse_number).astype(float)
    data["product"] = data["product"].str.split().str.join(" ")
    data["customer"] = data["customer"].str.split().str.join(" ")
    data.insert(0, "store", path.stem.replace("store_", "").replace("_", " ").title())

    reasons = data.apply(rejection_reason, axis=1)
    rejected = original[reasons != ""].copy()
    rejected.insert(0, "reason", reasons[reasons != ""])
    rejected.insert(0, "row", rejected.index + 1)  # row number in the original file
    rejected.insert(0, "file", path.name)

    valid = data[reasons == ""].copy()
    valid["quantity"] = valid["quantity"].astype(int)
    return valid, rejected


def unify_product_names(sales):
    """'WIRELESS MOUSE' and 'wireless mouse' are one product: the most common spelling wins."""
    key = sales["product"].str.casefold()
    preferred = sales.groupby(key)["product"].agg(lambda names: names.mode().iloc[0])
    sales["product"] = key.map(preferred)
    return sales


def style_sheet(sheet, money_columns=()):
    header_fill = PatternFill("solid", start_color="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    for index, column in enumerate(sheet.columns, start=1):
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
        sheet.column_dimensions[get_column_letter(index)].width = min(width + 3, 45)
        if column[0].value in money_columns:
            for cell in column[1:]:
                cell.number_format = MONEY_FORMAT
    sheet.freeze_panes = "A2"


def write_report(sales, rejected, output):
    sales = sales.sort_values(["date", "store", "product"])
    sales["revenue"] = (sales["quantity"] * sales["unit_price"]).round(2)

    monthly = sales.pivot_table(
        index=sales["date"].dt.strftime("%Y-%m").rename("month"),
        columns="store", values="revenue", aggfunc="sum", fill_value=0,
    )
    stores = list(monthly.columns)
    monthly["Total"] = monthly.sum(axis=1)

    products = (
        sales.groupby("product")
        .agg(units=("quantity", "sum"), revenue=("revenue", "sum"))
        .sort_values("revenue", ascending=False)
    )

    clean = sales.copy()
    clean["date"] = clean["date"].dt.date

    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        monthly.to_excel(writer, sheet_name="Monthly Revenue")
        products.to_excel(writer, sheet_name="Top Products")
        clean.to_excel(writer, sheet_name="Clean Data", index=False)
        rejected.to_excel(writer, sheet_name="Rejected Rows", index=False)

        style_sheet(writer.sheets["Monthly Revenue"], money_columns=stores + ["Total"])
        style_sheet(writer.sheets["Top Products"], money_columns=["revenue"])
        style_sheet(writer.sheets["Clean Data"], money_columns=["unit_price", "revenue"])
        style_sheet(writer.sheets["Rejected Rows"])

        sheet = writer.sheets["Monthly Revenue"]
        last_row = len(monthly) + 1
        chart = BarChart()
        chart.type = "col"
        chart.grouping = "stacked"
        chart.overlap = 100
        chart.title = "Revenue by month and store"
        chart.y_axis.title = "Revenue"
        chart.add_data(
            Reference(sheet, min_col=2, max_col=len(stores) + 1, min_row=1, max_row=last_row),
            titles_from_data=True,
        )
        chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=last_row))
        chart.height, chart.width = 9, 18
        sheet.add_chart(chart, f"{get_column_letter(len(stores) + 4)}2")

    return sales


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--input", type=Path, default=Path("sample_data"))
    parser.add_argument("--output", type=Path, default=Path("output/sales_report.xlsx"))
    args = parser.parse_args()

    files = sorted(
        path for path in args.input.iterdir()
        if path.suffix.lower() in {".csv", ".xlsx"} and not path.name.startswith("~$")
    )
    if not files:
        sys.exit(f"No .csv or .xlsx files found in {args.input}")

    valid_parts, rejected_parts, skipped = [], [], []
    for path in files:
        try:
            valid, rejected = load_store(path)
        except ValueError as error:
            skipped.append(f"{path.name}: {error}")
            continue
        valid_parts.append(valid)
        rejected_parts.append(rejected)
        print(f"{path.name:<22} {len(valid):>4} valid   {len(rejected):>2} rejected")

    if not valid_parts:
        sys.exit("No usable files:\n  " + "\n  ".join(skipped))

    sales = unify_product_names(pd.concat(valid_parts, ignore_index=True))
    before = len(sales)
    sales = sales.drop_duplicates()
    rejected = pd.concat(rejected_parts, ignore_index=True)

    sales = write_report(sales, rejected, args.output)

    print(f"\nDuplicates removed: {before - len(sales)}")
    print(f"Rows in report:     {len(sales)}")
    print(f"Total revenue:      {sales['revenue'].sum():,.2f}")
    for line in skipped:
        print(f"SKIPPED {line}")
    print(f"\nReport saved to {args.output}")


if __name__ == "__main__":
    main()
