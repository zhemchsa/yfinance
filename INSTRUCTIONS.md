# YFinance Workbook Automation Instructions

This project loads stock data from Yahoo Finance via `yfinance`, screens tickers, and writes the results into the workbook at `docs/hybrid_research_portfolio.xlsx`.

## 1. Project layout

- `src/main.py` — main Python entry point
- `src/__init__.py` — marks `src` as a Python package so modules can be imported cleanly
- `run_yfinance.sh` — shell runner for Unix/macOS users
- `requirements.txt` — Python dependencies
- `docs/hybrid_research_portfolio.xlsx` — Excel workbook used for outputs and manual overrides
- `.myenv3.14` — Python 3.14 virtual environment for this project

## 2. Create or activate the environment

From the project root:

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
```

If the environment is missing:

```bash
cd /Users/zhemchsa/repos/yfinance
python3 -m venv .myenv3.14
source .myenv3.14/bin/activate
pip install -r requirements.txt
```

## 3. Install dependencies

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
pip install -r requirements.txt
```

## 4. Run the automation

### Option A: Use the shell script

```bash
cd /Users/zhemchsa/repos/yfinance
./run_yfinance.sh
```

### Option B: Run Python directly

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
python src/main.py
```

### Option C: Run only the screen logic

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
python src/main.py --screen-only
```

### Option D: Run only the database refresh logic

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
python src/main.py --db-only
```

## 5. Open the workbook

Open the workbook here:

```text
docs/hybrid_research_portfolio.xlsx
```

If you are using Excel for Mac, Excel for Web, or another spreadsheet tool, open the workbook file directly. The script writes values into the `Data Import` section while preserving the original workbook formulas and formatting.

## 6. Where the script writes data

The script updates the workbook in the `Data Import` sheet.

### Ticker inputs
The workbook supports up to 30 reusable ticker inputs in row 23 of the `Data Import` sheet. Enter tickers in every other column, starting at B23 and ending at BH23 (B23, D23, F23, ... BH23). The columns between them are reserved for peer labels.

The script reads these ticker cells and refreshes up to 30 tickers.

### Output rows
The script writes rows beginning at the `Data Import` section for each ticker and metric such as:

- Share price
- Trailing P/E
- Market capitalization
- Shares outstanding
- Revenue
- EBIT
- Operating cash flow
- Capital expenditure
- Cash and equivalents
- Total debt
- Beta

For the Trailing P/E row, columns G-I show the ticker's sector, a weighted P/E benchmark based on Yahoo's top 50 sector companies, and the stock's premium or discount to that benchmark. The comparison is `(stock P/E / sector P/E) - 1`. Companies without a positive trailing P/E are excluded from the benchmark. This is a top-50 estimate, not a P/E calculated from every company in the sector. ETFs or funds without a Yahoo sector assignment may show `N/A` and have no sector comparison.

See the `PE Guide` workbook tab for a column-by-column explanation and an example.

The metric values are written in the `C` column (Yahoo value) and the formula in the `E` column uses the manual override logic:

```excel
=IF(D7="",C7,D7)
```

This means:
- if the yellow override cell `D7` is blank, the workbook uses the fetched value in `C7`
- if the yellow override cell `D7` has a manual entry, it overrides the fetched value

## 7. How to interpret the results in the workbook

### Dashboard tab
The `Dashboard` sheet summarizes the current investment case and portfolio snapshot.

Use the workbook to compare:
- case selected
- current price
- portfolio value
- weight assumptions
- valuation output values from related tabs

### Data Import tab
This is the main data refresh section.

Review the rows for each ticker and metric:
- `A` column: ticker
- `B` column: metric name
- `C` column: Yahoo Finance fetched value
- `D` column: manual override (yellow cell)
- `E` column: active value
- `F` column: refreshed timestamp

### Yellow fields
Yellow cells represent manual override inputs. You can type your own values there when you want to override the market-data default.

### Formula behavior
The workbook is designed so that manual overrides take precedence over fetched values. If you do not override a value, the workbook keeps the latest yfinance value.

## 8. What `__init__.py` does

`src/__init__.py` marks the `src` directory as a Python package.

This allows imports such as:

```python
from src.portfolio import Portfolio
```

or:

```python
from src.main import main
```

Without `__init__.py`, the folder is just a normal directory. Adding it makes the project easier to extend as you add more files like:

```text
src/
  __init__.py
  portfolio.py
  market_data.py
  main.py
```

Then the code can be organized cleanly instead of keeping everything in one big script.

## 9. Example package pattern

If you create a file named `src/portfolio.py` with a class or function, you would import it like this:

```python
from src.portfolio import Portfolio
```

or:

```python
from src.portfolio import build_portfolio
```

This is the normal package import style in Python projects.

## 10. Troubleshooting

### The workbook does not update
Check that:

```bash
source .myenv3.14/bin/activate
python src/main.py
```

is running without an error.

### The script fails on a ticker
Some tickers may not be valid or may return incomplete data. The script usually falls back to the next available value, but if a ticker is malformed, replace it in the workbook input cells.

### You see missing data
Open the workbook and confirm the corresponding `Data Import` rows are populated. If a metric is blank, check the ticker symbol and verify that Yahoo Finance returns data for it.

## 11. Recommended workflow

1. Open the workbook.
2. Update the reusable ticker input cells in `Data Import`.
3. Run the script.
4. Refresh the workbook file.
5. Review the metric rows and dashboard summary.
6. Use the yellow override cells only for manual adjustments.

## 12. Typical command sequence

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
python src/main.py
```

Then open:

```text
docs/hybrid_research_portfolio.xlsx
```

and review the updated values in the `Data Import` and `Dashboard` tabs.
