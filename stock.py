# Usage
# type "python stock.py" in command line
# type in the symbol of a stock (e.g. GME)
# or type "exit" to close the program 

import openpyxl
from prettytable import PrettyTable 

# load the excel file
excelFile = openpyxl.load_workbook('example.xlsx')	# load file
stockSheet = excelFile['transaction']	# load sheet

# testData = stockSheet.cell(column=2, row=2).value
# print(testData)

inputCmd = ''

while inputCmd != 'exit':
	inputCmd = input("Please enter the ticker symbol or enter 'exit' to close the program: ")

	if inputCmd == 'exit':
		break

	dateList = []
	actionList = []
	shareList = []
	priceList = []
	totalList = []
	GLList = []

	checkStockIndicator = 0		# check if the stock is in the file (default=0, found=1)
	rowCount = 2
	while True:
		checkSymbol = stockSheet.cell(column=3, row=rowCount).value
		if checkSymbol == None:
			if checkStockIndicator == 0:
				print("Unable to find '"+inputCmd+"'. Please enter in uppercase.\n")
				break
			else:
				break
		if inputCmd == checkSymbol:
			checkStockIndicator = 1
		rowCount += 1

	if checkStockIndicator == 1:
		stockSym = inputCmd
		rowCount = 2
		actionCount = 0
		while True:

			if stockSheet.cell(column=3, row=rowCount).value == None:
				rowCount -= 1
				break

			if stockSheet.cell(column=3, row=rowCount).value == stockSym:
				dateStr = str(stockSheet.cell(column=1, row=rowCount).value)
				dateList.append(dateStr[:-9])
				actionList.append(stockSheet.cell(column=2, row=rowCount).value)
				share = int(stockSheet.cell(column=4, row=rowCount).value)
				shareList.append(share)
				price = round(float(stockSheet.cell(column=5, row=rowCount).value), 2)
				priceList.append(price)
				total = share * price
				totalList.append(total)
				actionCount += 1
			rowCount += 1

		stockTable = PrettyTable(['Date', 'Action', 'Share', 'Price', 'Total'])
		for r in range(0, actionCount):
			stockTable.add_row([dateList[r], actionList[r], shareList[r], priceList[r], totalList[r]])

		print("\nStock:", stockSym)
		print(stockTable, "\n\n")

	# print(dateList)
	# print(actionList)
	# print(shareList)
	# print(priceList)


