# Automated Excel Report Generator

A Python command-line tool that reads a CSV file, cleans it, and writes a formatted Excel report with summary KPIs, aggregation tables, and charts.

**Stack:** Python, pandas, NumPy, openpyxl, matplotlib
**Tested on:** NYC Rolling Sales (public NYC Department of Finance data, via Kaggle), [N] records

## Usage

```
pip install -r requirements.txt
python main.py path/to/your_file.csv
python main.py path/to/your_file.csv --output My_Report.xlsx --verbose
```

## What the report contains

| Sheet | Contents |
|---|---|
| Executive Summary | Record count, date range, total and average sales, up to 10 generated insights |
| Cleaned Data | The dataset after cleaning (not the original file) |
| Monthly Analysis | Sales, transaction count, and average by calendar month (chronological) |
| Product Analysis | Top 20 categories by total value |
| Regional Analysis | Totals by region (borough, neighborhood, or ZIP, depending on the file) |
| Charts | Four static PNG charts: monthly trend, top categories, regional share, summary |

## How it works

1. **Load:** tries UTF-8, then Latin-1, then CP1252.
2. **Standardize columns:** lower-cases and renames known column names exactly (for example `sale_date` to `date`); drops leftover index columns such as `Unnamed: 0`.
3. **Clean:** converts placeholder values (`-`, blanks, `N/A`) to missing, converts price strings to numbers, fills missing text with "Unknown", removes duplicate rows. Missing numbers stay missing; they are not filled with a median or mode unless `DataProcessor(impute=True)` is set.
4. **Detect domain:** keyword matching on column names scores six domains (sales, real estate, financial, HR, marketing, inventory). This is rule-based, not machine learning.
5. **Aggregate and report:** groups valid sales (price above 0) by month, category, and region, then writes the workbook.

## Insights generated

- **Sales and real estate files:** top product or property type, revenue concentration, median price, most active borough or neighborhood, monthly trend.
- **All files:** share of missing values and most incomplete columns (measured on the source file, before cleaning), duplicate count, heavily skewed numeric columns.
- **Financial, HR, marketing, inventory:** detected, but they receive only the general insights above. Domain-specific logic is not implemented.

## Limitations

- Column mapping is heuristic and was developed against NYC Rolling Sales and simple sales files. Other layouts may need mapping changes in `standardize_columns`.
- Charts are static images, not interactive.
- The "Product" and "Region" labels in sheets are generic; for NYC data they are building class and borough.
- No automated tests yet.

## Project structure

```
├── main.py                 # command-line entry point
├── data_processor.py       # loading, cleaning, aggregation
├── data_intelligence.py    # domain detection and insights
├── report_generator.py     # Excel workbook and formatting
├── chart_generator.py      # matplotlib charts
├── config.py               # settings (partly unused)
└── requirements.txt
```

## Sample output

[Paste the Key Insights printed in the Executive Summary from a run on your real NYC file, and add a screenshot of the report.]
