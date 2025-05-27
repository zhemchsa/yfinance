import yfinance as yf

tickers = ['AAPL', 'GE', 'PLTR']

metrics = [
    'trailingPE',
    'pegRatio',
    'priceToBook',
    'debtToEquity',
    'returnOnEquity',
    'freeCashflow'
]

print(f"{'Ticker':<6} {'P/E':>8} {'PEG':>8} {'P/B':>8} {'Debt/Equity':>14} {'ROE':>10} {'Free Cash Flow ($)':>22}")
print("=" * 80)

for ticker in tickers:
    stock = yf.Ticker(ticker)
    info = stock.info
    
    values = []
    for metric in metrics:
        value = info.get(metric)
        if isinstance(value, (int, float)):
            if metric == 'returnOnEquity':
                value = f"{value*100:.2f}%"  # Convert ROE to %
            elif metric == 'freeCashflow':
                value = f"${value/1e9:.2f}B"  # Convert FCF to billions
            else:
                value = f"{value:.2f}"
        else:
            value = "N/A"
        values.append(value)
    
    print(f"{ticker:<6} {values[0]:>8} {values[1]:>8} {values[2]:>8} {values[3]:>14} {values[4]:>10} {values[5]:>22}")