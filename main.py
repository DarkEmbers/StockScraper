import time
import yfinance as yf
import threading
from openpyxl import Workbook, worksheet, load_workbook
import json
import os
import sys
import schedule

workbook = None
sheet = None
sheetPath = ""
threads = []


def get_stats(ticker, row_count):
	# get the stock data
	info = yf.Tickers(ticker).tickers[ticker].info
	update_stock_price(info['currentPrice'], row_count)
	# print(f"{ticker}::{info['currentPrice']}")


def update_stock_price(price, row_count):
	# write the stock price to the cell values of column B
	sheet.cell(row=row_count, column=2).value = price
	workbook.save(sheetPath)


def main():
	global workbook, sheet, sheetPath

	try:
		# load config.json and get excel sheet path
		with open(config_path) as f:
			config = json.load(f)
			sheetPath = config["path"]

	except FileNotFoundError:
		print("Config file not found")
		return

	try:
		workbook = load_workbook(sheetPath)
		sheet = workbook.active

	except FileNotFoundError:
		print("Excel file not found")
		return

	# get the ticker symbols from the first column of the spreadsheet
	row_count = 2
	while True:
		# Read the cell values of column A
		ticker = sheet.cell(row=row_count, column=1).value
		if ticker is None:
			break

		# create a new thread for each ticker symbol
		thread = threading.Thread(target=get_stats, args=[ticker, row_count])
		thread.start()
		threads.append(thread)
		row_count += 1


if __name__ == '__main__':
	# determine if the application is a .exe
	if getattr(sys, 'frozen', False):
		application_path = os.path.dirname(sys.executable)
	# or a script file
	elif __file__:
		application_path = os.path.dirname(__file__)
		config_name = 'StonksConfig.json'
		config_path = os.path.join(application_path, config_name)

	main()
	for thread in threads:
		thread.join()

	print(f"Updated {time.clock_gettime(time.CLOCK_REALTIME)}")

	# schedule.every(15).minutes.do(main)
	# while True:
	# 	schedule.run_pending()
	# 	time.sleep(1)
