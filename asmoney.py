import openpyxl
import os

# Files must be named "*_Transactions.xlsx"
# Workbooks should have all transactions in a sheet named "Journal"


#os.getcwd()
#os.chdir()

# Handle inputs


wb = openpyxl.load_workbook('example.xlsx')

sheet = wb.get_sheet_by_name('Sheet1')

max_column = sheet.max_column
max_row = sheet.max_row

for cellObj in sheet.columns[1]:
    print(cellObj.value)

for rowNum in range(2, sheet.max_row):  # skip the first row


#wb.save('example.xlsx')

#sheet['A1'] = 'Hello World!'

#sheet.freeze_panes = 'A2'

# Search input directory for files *_Transactions.xlsx


# Open each workbook and select the "Journal" tab

