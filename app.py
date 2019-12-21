import openpyxl
import os

def main():
    # Insert the file of the spreadsheet you want to change here
    fileName = 'december90DaysTransactions.xlsx'

    # Make sure to go into the Excel sheet and change the sheet name to transactions prior to running the program
    # Sheet name you would like to change
    sheetName = 'transactions'
    # Load the workbook
    wb = openpyxl.load_workbook(fileName)
    
    # Access the sheet in the workbook
    transactionsWorksheet = wb[sheetName]

    # If E1 is tags, then delete column 5
    delete_tags_column(transactionsWorksheet)
    wb.save(fileName)

    organizeExpenses(transactionsWorksheet)


def delete_tags_column(transactionsWorksheet):
    tagsValue = transactionsWorksheet['E1'].value
    if (tagsValue == 'Tags'):
        transactionsWorksheet.delete_cols(5)

def organizeExpenses(transactionsWorksheet):
    # Maximum row of the sheet
    maxRow = transactionsWorksheet.max_row

    transactionCategories = {}
    
    for i in range(2, maxRow + 1):
        category = transactionsWorksheet.cell(row=i, column=4).value
        transaction = transactionsWorksheet.cell(row = i, column=5).value
        # print(category, '$', transaction)

        if category in transactionCategories:
            transactionCategories[category] += transaction
        elif category not in transactionCategories:
            transactionCategories[category] = 0
        
    
    for key in transactionCategories:
        print(key, '$', round(transactionCategories[key], 2))


if __name__ == "__main__":
    main()

