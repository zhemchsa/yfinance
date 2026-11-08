import yfinance as yf

# Example tickers - you can load more or use an index
tickers = ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'TSLA', 'JPM', 'WMT', 'INTC', 'CVX']

for ticker in tickers:
    stock = yf.Ticker(ticker)
    info = stock.info

    pe_ratio = info.get('trailingPE')
    pb_ratio = info.get('priceToBook')
    peg_ratio = info.get('pegRatio')

    if pe_ratio and pb_ratio and peg_ratio:
        if pe_ratio < 15 and pb_ratio < 1.5 and peg_ratio < 1:
            print(f"{ticker}: P/E={pe_ratio}, P/B={pb_ratio}, PEG={peg_ratio}")

            