# YFinance Portfolio Workbook Automation

This project uses Python and `yfinance` to collect stock data, screen tickers, and write selected results into the workbook at `docs/hybrid_research_portfolio.xlsx`.

## Project structure

- `src/main.py` — main script that loads tickers, fetches data, and writes workbook values
- `run_yfinance.sh` — macOS/Linux runner script
- `requirements.txt` — project dependencies
- `docs/hybrid_research_portfolio.xlsx` — Excel workbook used for data review and manual overrides
- `INSTRUCTIONS.md` — detailed execution notes

## Prerequisites

- Python 3.12+ recommended
- Internet access for Yahoo Finance data
- A workbook file at `docs/hybrid_research_portfolio.xlsx`

## 1. macOS / Linux setup

From the project root:

```bash
cd /Users/zhemchsa/repos/yfinance
python3 -m venv .myenv3.14
source .myenv3.14/bin/activate
pip install -r requirements.txt
```

Run the project:

```bash
cd /Users/zhemchsa/repos/yfinance
source .myenv3.14/bin/activate
python src/main.py
```

Or use the shell runner:

```bash
cd /Users/zhemchsa/repos/yfinance
./run_yfinance.sh
```

Screen-only mode:

```bash
cd /Users/zhemchsa/repos/yfinance
./run_yfinance.sh --screen-only
```

Database-only mode:

```bash
cd /Users/zhemchsa/repos/yfinance
./run_yfinance.sh --db-only
```

## 2. Windows setup

From Command Prompt:

```cmd
cd C:\path\to\yfinance
python -m venv .myenv3.14
.\.myenv3.14\Scripts\activate
pip install -r requirements.txt
```

Run the script:

```cmd
cd C:\path\to\yfinance
.\.myenv3.14\Scripts\activate
python src\main.py
```

Or use PowerShell:

```powershell
cd C:\path\to\yfinance
.\.myenv3.14\Scripts\Activate.ps1
python .\src\main.py
```

## 3. Open the workbook

Open this file:

```text
docs/hybrid_research_portfolio.xlsx
```

The workbook has a `Data Import` section where values are populated from Yahoo Finance.

## 4. How the workbook works

The workbook uses a `Data Import` section with these common patterns:

- `A` column: ticker
- `B` column: metric name
- `C` column: Yahoo value fetched by Python
- `D` column: manual override (yellow cells)
- `E` column: active value used by formulas
- `F` column: refresh timestamp

If the manual override cell is blank, the workbook uses the fetched value. If the yellow cell contains a value, it overrides the fetched value.

## 5. What the script updates

The script reads the ticker values from the workbook and then writes refreshed values around the `Data Import` area for each ticker and metric. This keeps the workbook structure intact while refreshing the data.

## 6. Interpretation of results

- Use the `Dashboard` tab to review the portfolio summary and case-level output.
- Use the `Data Import` tab to review raw ticker metrics.
- Use yellow override cells only for manual adjustments.
- Keep the `C` values as fetched data and `D` values as user overrides.

## 7. Troubleshooting

### Worksheet not updating
Make sure the script runs without errors:

```bash
python src/main.py
```

### Missing value
Check the ticker symbol and confirm the data exists in Yahoo Finance.

### Manual values not being used
Confirm the override cell in the `D` column is populated and the formula in the `E` column is still active.

## 8. Next steps

- Update the workbook ticker cells in the reusable input range.
- Run the project script.
- Review the refreshed values in the workbook.
- Use the yellow override cells when you want to manually replace any fetched number.
