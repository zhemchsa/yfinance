from dotenv import load_dotenv
import yfinance as yf
import psycopg2
from decimal import Decimal
import os

load_dotenv()

DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_HOST = os.getenv("DB_HOST")
DB_PORT = os.getenv("DB_PORT")


# Add more tickers from the same industry as GE (Industrials/Conglomerates/Industrial Products)
tickers = [
    'GE',      # General Electric
    'HON',     # Honeywell
    'MMM',     # 3M
    'EMR',     # Emerson Electric
    'DHR',     # Danaher
    'ETN',     # Eaton
    'ROK',     # Rockwell Automation
    'PH',      # Parker-Hannifin
    'ITW',     # Illinois Tool Works
    'IR',      # Ingersoll Rand
    'XYL',     # Xylem
    'AOS',     # A. O. Smith
    'GWW',     # W.W. Grainger
    'JCI',     # Johnson Controls
    'PNR',     # Pentair
    'SWK',     # Stanley Black & Decker
    'TT',      # Trane Technologies
    'ALC',     # Alcon
    'AME',     # AMETEK
    'AVY',     # Avery Dennison
]

metrics = [
    'trailingPE',
    'pegRatio',
    'priceToBook',
    'debtToEquity',
    'returnOnEquity',
    'freeCashflow'
]

# Connect to PostgreSQL
conn = psycopg2.connect(
    dbname=DB_NAME,
    user=DB_USER,
    password=DB_PASSWORD,
    host=DB_HOST,
    port=DB_PORT
)
cur = conn.cursor()

# Create table if not exists
cur.execute("""
CREATE TABLE IF NOT EXISTS stock_metrics (
    ticker TEXT PRIMARY KEY,
    trailing_pe NUMERIC,
    peg_ratio NUMERIC,
    price_to_book NUMERIC,
    debt_to_equity NUMERIC,
    return_on_equity NUMERIC,
    free_cashflow NUMERIC
)
""")
cur.execute("ALTER TABLE stock_metrics ADD COLUMN IF NOT EXISTS industry TEXT;")
conn.commit()

# Delete existing data
cur.execute("DELETE FROM stock_metrics;")
conn.commit()

# Insert or update data
for ticker in tickers:
    stock = yf.Ticker(ticker)
    info = stock.info

    values = []
    for metric in metrics:
        value = info.get(metric)
        if isinstance(value, (int, float)):
            values.append(value)
        else:
            values.append(None)
    industry = info.get("industry", "N/A")

    cur.execute("""
        INSERT INTO stock_metrics (ticker, trailing_pe, peg_ratio, price_to_book, debt_to_equity, return_on_equity, free_cashflow, industry)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        ON CONFLICT (ticker) DO UPDATE SET
            trailing_pe = EXCLUDED.trailing_pe,
            peg_ratio = EXCLUDED.peg_ratio,
            price_to_book = EXCLUDED.price_to_book,
            debt_to_equity = EXCLUDED.debt_to_equity,
            return_on_equity = EXCLUDED.return_on_equity,
            free_cashflow = EXCLUDED.free_cashflow,
            industry = EXCLUDED.industry
    """, (ticker, *values, industry))
    conn.commit()

# Query and print per-ticker results with industry P/E aggregate
cur.execute(f"""
SELECT
    s.ticker,
    s.trailing_pe,
    s.peg_ratio,
    s.price_to_book,
    s.debt_to_equity,
    s.return_on_equity,
    s.free_cashflow,
    s.industry,
    i.industry_pe
FROM stock_metrics s
LEFT JOIN (
    SELECT industry, AVG(trailing_pe) AS industry_pe
    FROM stock_metrics
    WHERE trailing_pe IS NOT NULL
    GROUP BY industry
) i ON s.industry = i.industry
WHERE s.ticker IN ({','.join(['%s']*len(tickers))})
ORDER BY s.ticker
""", tickers)
rows = cur.fetchall()

print("\nPer-Ticker Metrics:")
header = "| {:^8} | {:^8} | {:^18} | {:^8} | {:^8} | {:^14} | {:^10} | {:^18} | {:^25} |".format(
    "Ticker", "P/E", "Industry Avg P/E", "PEG", "P/B", "Debt/Equity", "ROE", "Free Cash Flow ($B)", "Industry"
)
separator = "+" + "-"*10 + "+" + "-"*10 + "+" + "-"*20 + "+" + "-"*10 + "+" + "-"*10 + "+" + "-"*16 + "+" + "-"*12 + "+" + "-"*20 + "+" + "-"*27 + "+"
print(separator)
print(header)
print(separator)
for row in rows:
    ticker = row[0]
    pe = f"{row[1]:.2f}" if row[1] is not None else "N/A"
    industry_pe = f"{row[-1]:.2f}" if row[-1] is not None else "N/A"
    formatted = []
    for idx, val in enumerate(row[2:-2]):  # skip ticker, P/E, industry_pe
        if val is None:
            formatted.append("N/A")
        elif idx == 4:  # Free Cash Flow column (now index 4 in this slice)
            formatted.append(f"${(val/Decimal('1e9')):.2f}B")
        else:
            formatted.append(f"{val:.2f}")
    industry = row[-2] if row[-2] else "N/A"
    print("| {:^8} | {:^8} | {:^18} | {:^8} | {:^8} | {:^14} | {:^10} | {:^18} | {:^25} |".format(
        ticker, pe, industry_pe, *formatted, industry
    ))
print(separator)

cur.close()
conn.close()