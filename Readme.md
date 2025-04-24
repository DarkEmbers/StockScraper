# Stocks Auto-Updater
A Python-based tool that automatically updates stock prices in an Excel spreadsheet using data from TradingView's technical analysis API.

## Requirementes
- openpyxl
- tradingview_ta
- schedule

```bash
pip install openpyxl tradingview_ta schedule
```

## Excel File Format
| A (Ticker) | B (Country) | C (Exchange) | D (Price) |
|------------|-------------|--------------|-----------|
| AAPL       | america     | NASDAQ       |           |
| TSLA       | america     | NASDAQ       |           |

- Column A: Ticker symbol (e.g., `AAPL`, `TSLA`)
- Column B: Screener/country for TradingView (e.g., `america`)
- Column C: Stock exchange (e.g., `NASDAQ`)
- Column D: This column will be automatically updated with the latest price

## Config File
This JSON file includes:
- path: Path to the Excel sheet.
- interval: Update frequency in minutes.

## Usage
```bash
python main.py
```

