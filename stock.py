# Usage
# type "python stock.py" in command line
# type in the symbol of a stock (e.g. GME)
# or type "exit" to close the 

import openpyxl
from prettytable import PrettyTable 

# load the excel file
excelFile = openpyxl.load_workbook('example.xlsx')	# load file
stockSheet = excelFile['transaction']	# load sheet

testData = stockSheet.cell(column=2, row=2).value
print(testData)
