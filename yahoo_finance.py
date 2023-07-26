### import packages

import yfinance as yf
from openpyxl import load_workbook

### import packages

import yfinance as yf
from openpyxl import load_workbook

### loading workbook

wb = load_workbook('yahoo.xlsx')
ws = wb.active

### get ticker list from the first row

ticker_list = []

for row in ws.iter_rows(min_row=1, max_row=1, min_col=2, max_col=len(ws['1'])):
    for cell in row:
        ticker_list.append(cell.value)

### load tickers

loaded_tickers = []

for ticker in ticker_list:
    ticker = yf.Ticker(ticker)
    loaded_tickers.append(ticker)

### get data for each ticker

all_ticker_data = []

for i in range(0, len(loaded_tickers)):
    all_ticker_data.extend((loaded_tickers[i].info.get('currentPrice', 'N/A'), loaded_tickers[i].info.get('trailingPE', 'N/A'), loaded_tickers[i].info.get('forwardPE', 'N/A'), loaded_tickers[i].info.get('trailingEps', 'N/A'),
                        loaded_tickers[i].info.get('beta', 'N/A'), loaded_tickers[i].info.get('fiftyTwoWeekHigh', 'N/A'), loaded_tickers[i].info.get('fiftyTwoWeekLow', 'N/A'), loaded_tickers[i].info.get('fiftyDayAverage', 'N/A'),
                        loaded_tickers[i].info.get('twoHundredDayAverage', 'N/A'), loaded_tickers[i].info.get('dividendRate', 'N/A'), loaded_tickers[i].info.get('dividendYield', 'N/A')))
    
### put data into worksheet

i = 0

for col in ws.iter_cols(min_row=2, max_row=len(ws['A']), min_col=2, max_col=len(ws['1'])):
    for cell in col:
        cell.value = all_ticker_data[i]
        i += 1

### save workbook
wb.save('yahoo_updated.xlsx')

