#!/usr/bin/env python3

import math
import os
import sys
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path

import yfinance as yf
from dotenv import load_dotenv
import psycopg2
from openpyxl import load_workbook


DEFAULT_TICKERS = [
    'GE', 'TSLA', 'PLTR', 'AMZN', 'HOOD', 'NVDA',
    'QQQ', 'VOO', 'IBIT', 'TEM', 'CINF', 'FETH',
    'BITX', 'BMNR', 'SPCX', 'AMD']
MAX_TICKERS = 30
TICKER_INPUT_COLUMNS = tuple(2 + (2 * index) for index in range(MAX_TICKERS))

METRICS = [
    'trailingPE',
    'pegRatio',
    'priceToBook',
    'debtToEquity',
    'returnOnEquity',
    'freeCashflow'
]

WORKBOOK_PATH = Path(__file__).resolve().parent.parent / 'docs' / 'hybrid_research_portfolio.xlsx'
DATA_IMPORT_METRICS = [
    'Share price',
    'Market capitalization',
    'Shares outstanding',
    'Revenue',
    'EBIT',
    'Operating cash flow',
    'Capital expenditure',
    'Cash and equivalents',
    'Total debt',
    'Beta',
    'Trailing P/E',
]


def screen_value_candidates(tickers):
    print(f"{'Ticker':<6} {'P/E':>8} {'P/B':>8} {'PEG':>8}")
    print("-" * 36)

    matches = []
    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.info
        pe_ratio = info.get('trailingPE')
        pb_ratio = info.get('priceToBook')
        peg_ratio = info.get('pegRatio')

        if pe_ratio is not None and pb_ratio is not None and peg_ratio is not None:
            if pe_ratio < 15 and pb_ratio < 1.5 and peg_ratio < 1:
                matches.append((ticker, pe_ratio, pb_ratio, peg_ratio))

    if not matches:
        print("No tickers matched the screen.")
        return

    for ticker, pe_ratio, pb_ratio, peg_ratio in matches:
        print(f"{ticker:<6} {pe_ratio:>8.2f} {pb_ratio:>8.2f} {peg_ratio:>8.2f}")


def print_stock_metrics_table(tickers):
    print(f"{'Ticker':<6} {'P/E':>8} {'PEG':>8} {'P/B':>8} {'Debt/Equity':>14} {'ROE':>10} {'Free Cash Flow ($)':>22}")
    print("=" * 80)

    for ticker in tickers:
        stock = yf.Ticker(ticker)
        info = stock.info

        values = []
        for metric in METRICS:
            value = info.get(metric)
            if isinstance(value, (int, float)):
                if metric == 'returnOnEquity':
                    value = f"{value * 100:.2f}%"
                elif metric == 'freeCashflow':
                    value = f"${value / 1e9:.2f}B"
                else:
                    value = f"{value:.2f}"
            else:
                value = 'N/A'
            values.append(value)

        print(f"{ticker:<6} {values[0]:>8} {values[1]:>8} {values[2]:>8} {values[3]:>14} {values[4]:>10} {values[5]:>22}")


def get_db_connection():
    load_dotenv()

    db_name = os.getenv('DB_NAME')
    db_user = os.getenv('DB_USER')
    db_password = os.getenv('DB_PASSWORD')
    db_host = os.getenv('DB_HOST', 'localhost')
    db_port = os.getenv('DB_PORT', '5432')

    if not all([db_name, db_user, db_password]):
        print('Skipping PostgreSQL refresh: database environment variables are not configured.')
        return None

    return psycopg2.connect(
        dbname=db_name,
        user=db_user,
        password=db_password,
        host=db_host,
        port=db_port,
    )


def get_workbook_tickers():
    if not WORKBOOK_PATH.exists():
        return []

    workbook = load_workbook(WORKBOOK_PATH, data_only=False)
    if 'Data Import' not in workbook.sheetnames:
        workbook.close()
        return []

    try:
        sheet = workbook['Data Import']
        values = []
        for column in TICKER_INPUT_COLUMNS:
            value = sheet.cell(row=23, column=column).value
            if value is None:
                continue
            text = str(value).strip()
            if text and text not in {'None', 'Peer 1', 'Peer 2', 'Peer 3'}:
                values.append(text)
        return values[:MAX_TICKERS]
    finally:
        workbook.close()


def get_sector_average_pe(sector_name):
    if not isinstance(sector_name, str) or not sector_name.strip():
        return None

    normalized_name = sector_name.strip().casefold()
    if normalized_name in {'n/a', 'none'}:
        return None

    sector_key = normalized_name.replace(' ', '-')
    top_companies = yf.Sector(sector_key).top_companies
    if top_companies is None:
        return None

    total_weight = 0.0
    weighted_earnings_yield = 0.0
    for symbol, company in top_companies.iterrows():
        weight = company['market weight']
        if not isinstance(weight, (int, float)) or not math.isfinite(weight) or weight <= 0:
            continue

        trailing_pe = yf.Ticker(symbol).info.get('trailingPE')
        if not isinstance(trailing_pe, (int, float)) or not math.isfinite(trailing_pe) or trailing_pe <= 0:
            continue

        total_weight += weight
        weighted_earnings_yield += weight / trailing_pe

    if weighted_earnings_yield == 0:
        return None
    return total_weight / weighted_earnings_yield


def refresh_workbook_data(tickers):
    if not WORKBOOK_PATH.exists():
        return

    workbook = load_workbook(WORKBOOK_PATH, data_only=False)
    if 'Data Import' not in workbook.sheetnames:
        workbook.close()
        return

    sheet = workbook['Data Import']
    timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')
    sheet['G6'] = 'Sector'
    sheet['H6'] = 'Sector P/E (Top 50)'
    sheet['I6'] = 'P/E vs Sector (%)'
    sheet.column_dimensions['I'].width = 20

    tickers = tickers[:MAX_TICKERS]
    for column in TICKER_INPUT_COLUMNS:
        sheet.cell(row=23, column=column).value = None

    for idx, ticker in enumerate(tickers):
        sheet.cell(row=23, column=TICKER_INPUT_COLUMNS[idx]).value = ticker

    sector_pe_cache = {}
    for idx, ticker in enumerate(tickers):
        info = yf.Ticker(ticker).info
        sector_name = info.get('sector')
        sector_label = sector_name if isinstance(sector_name, str) and sector_name.strip() else 'N/A'
        sector_key = sector_label.strip().casefold().replace(' ', '-')
        if sector_label != 'N/A' and sector_key not in sector_pe_cache:
            sector_pe_cache[sector_key] = get_sector_average_pe(sector_label)
        sector_pe = sector_pe_cache.get(sector_key)
        row_start = 7 + (idx * 19)

        metric_map = {
            'Share price': info.get('currentPrice'),
            'Market capitalization': info.get('marketCap'),
            'Shares outstanding': info.get('sharesOutstanding'),
            'Revenue': info.get('totalRevenue'),
            'EBIT': info.get('ebit') or info.get('operatingIncome'),
            'Operating cash flow': info.get('operatingCashflow'),
            'Capital expenditure': info.get('capitalExpenditures'),
            'Cash and equivalents': info.get('totalCash') or info.get('totalCash') or info.get('cash'),
            'Total debt': info.get('totalDebt'),
            'Beta': info.get('beta'),
            'Trailing P/E': info.get('trailingPE'),
        }

        for metric_index, metric_name in enumerate(DATA_IMPORT_METRICS):
            row = row_start + metric_index
            sheet.cell(row=row, column=1, value=ticker)
            sheet.cell(row=row, column=2, value=metric_name)
            sheet.cell(row=row, column=3, value=metric_map.get(metric_name))
            sheet.cell(row=row, column=4, value=None)
            sheet.cell(row=row, column=5, value=f'=IF(D{row}="",C{row},D{row})')
            sheet.cell(row=row, column=6, value=timestamp)
            for column in (7, 8, 9):
                sheet.cell(row=row, column=column).value = None
            if metric_name == 'Trailing P/E':
                sheet.cell(row=row, column=3).number_format = '0.00'
                sheet.cell(row=row, column=7, value=sector_label)
                sheet.cell(row=row, column=8, value=sector_pe)
                sheet.cell(
                    row=row,
                    column=9,
                    value=f'=IF(AND(ISNUMBER(E{row}),E{row}>0,ISNUMBER(H{row}),H{row}>0),E{row}/H{row}-1,"")',
                )
                sheet.cell(row=row, column=8).number_format = '0.00'
                sheet.cell(row=row, column=9).number_format = '0.0%'

    workbook.save(WORKBOOK_PATH)
    workbook.close()


def refresh_stock_metrics(tickers):
    conn = get_db_connection()
    if conn is None:
        return

    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS stock_metrics (
                    ticker TEXT PRIMARY KEY,
                    trailing_pe NUMERIC,
                    peg_ratio NUMERIC,
                    price_to_book NUMERIC,
                    debt_to_equity NUMERIC,
                    return_on_equity NUMERIC,
                    free_cashflow NUMERIC,
                    industry TEXT
                )
            """)
            conn.commit()

            cur.execute('DELETE FROM stock_metrics;')
            conn.commit()

            for ticker in tickers:
                stock = yf.Ticker(ticker)
                info = stock.info
                values = []
                for metric in METRICS:
                    value = info.get(metric)
                    if isinstance(value, (int, float)):
                        values.append(value)
                    else:
                        values.append(None)
                industry = info.get('industry', 'N/A')

                cur.execute(
                    """
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
                    """,
                    (ticker, *values, industry)
                )
                conn.commit()

            cur.execute(
                """
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
                WHERE s.ticker IN ({})
                ORDER BY s.ticker
                """.format(','.join(['%s'] * len(tickers))),
                tickers,
            )
            rows = cur.fetchall()

        print("\nPer-Ticker Metrics:")
        header = "| {:^8} | {:^8} | {:^18} | {:^8} | {:^8} | {:^14} | {:^10} | {:^18} | {:^25} |".format(
            'Ticker', 'P/E', 'Industry Avg P/E', 'PEG', 'P/B', 'Debt/Equity', 'ROE', 'Free Cash Flow ($B)', 'Industry'
        )
        separator = '+' + '-' * 10 + '+' + '-' * 10 + '+' + '-' * 20 + '+' + '-' * 10 + '+' + '-' * 10 + '+' + '-' * 16 + '+' + '-' * 12 + '+' + '-' * 20 + '+' + '-' * 27 + '+'
        print(separator)
        print(header)
        print(separator)

        for row in rows:
            ticker = row[0]
            pe = f"{row[1]:.2f}" if row[1] is not None else 'N/A'
            industry_pe = f"{row[-1]:.2f}" if row[-1] is not None else 'N/A'
            formatted = []
            for idx, val in enumerate(row[2:-2]):
                if val is None:
                    formatted.append('N/A')
                elif idx == 4:
                    formatted.append(f"${(Decimal(str(val)) / Decimal('1e9')):.2f}B")
                else:
                    formatted.append(f"{float(val):.2f}")
            industry = row[-2] if row[-2] else 'N/A'
            print("| {:^8} | {:^8} | {:^18} | {:^8} | {:^8} | {:^14} | {:^10} | {:^18} | {:^25} |".format(
                ticker, pe, industry_pe, *formatted, industry
            ))
        print(separator)
    finally:
        conn.close()


def main():
    args = sys.argv[1:]
    tickers = (get_workbook_tickers() or DEFAULT_TICKERS)[:MAX_TICKERS]

    if '--screen-only' in args:
        screen_value_candidates(tickers)
        refresh_workbook_data(tickers)
        return
    if '--db-only' in args:
        refresh_stock_metrics(tickers)
        refresh_workbook_data(tickers)
        return

    screen_value_candidates(tickers)
    print_stock_metrics_table(tickers)
    refresh_workbook_data(tickers)
    refresh_stock_metrics(tickers)


if __name__ == '__main__':
    main()
