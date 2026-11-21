# Sales Report Automation

A Python script that turns messy sales exports from several stores into one clean Excel report.

> Demonstration project built with fictional data.

## The problem

A retailer with three stores gets a sales export from each one every month. Each store's system exports differently:

| Store | File | Quirks |
|---|---|---|
| North | CSV, comma-separated | ISO dates (`2026-04-05`), dot decimals |
| South | CSV, semicolon-separated | Day-first dates (`07/04/2026`), comma decimals (`1.249,00`), stray blank lines |
| East | Excel | Title rows above the header, text dates (`Apr 07, 2026`), prices typed as text (`$1,249.00`) |

On top of that, the files contain duplicated rows, product names typed in different ways (`WIRELESS MOUSE`, `wireless mouse`, `  Wireless Mouse `) and a few broken values. Someone spends hours every month copying and fixing this by hand.

## The solution

One command reads every file in a folder and produces the report:

```
python consolidate.py --input sample_data --output output/sales_report.xlsx
```

```
store_east.xlsx         100 valid    3 rejected
store_north.csv         122 valid    2 rejected
store_south.csv         112 valid    1 rejected

Duplicates removed: 9
Rows in report:     325
Total revenue:      234,956.00
```

The script works out each file's layout by itself:

- **Delimiter and header position** are detected, so title rows and blank lines are not a problem.
- **Column names** are matched through a list of known aliases (`Qty`, `quantity`, `UNITS` all mean quantity).
- **Dates** are parsed with the format that fits the most rows in that file.
- **Numbers** are read whether they use `1,249.00` or `1.249,00`, with or without a currency symbol.
- **Product names** that differ only in case or spacing are merged into the most common spelling.
- **Duplicated rows** are removed.

Rows that cannot be trusted are never silently dropped or guessed. They go to a separate sheet with the file name, the row number and the reason, so a person can fix them at the source.

## The report

`output/sales_report.xlsx` has four sheets:

| Sheet | Content |
|---|---|
| Monthly Revenue | Revenue per month and store, with totals and a chart |
| Top Products | Units and revenue per product, best sellers first |
| Clean Data | Every valid sale in one consistent table |
| Rejected Rows | Rows left out, with file, row number and reason |

## Running it

Requires Python 3.10 or newer.

```
pip install -r requirements.txt
python make_sample_data.py
python consolidate.py
```

`make_sample_data.py` recreates the three messy sample files. To use your own data, put the `.csv` and `.xlsx` files in a folder and pass it with `--input`.

## Adapting it

- **New store**: drop its file into the input folder. If it uses a column name not seen before, add it to `ALIASES` at the top of `consolidate.py`.
- **New date format**: add it to `DATE_FORMATS`.
- **Assumption to review**: two rows identical in every field are treated as an export error and counted once. If your business can have genuine identical sales on the same day, include an order ID column in the exports.
