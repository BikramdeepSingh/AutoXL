import openpyxl as xl


def process_spreadsheet(filename):
    wb = xl.load_workbook(filename)
    sheet = wb['Sheet1']

    #to get the coordinate of a cell=sheet.cell(x,y) or sheet[a1] 
    #and cell.value to fetch the value in a cell

    for row in range(2,sheet.max_row+1):
        cell_value = sheet.cell(row,3)
        corrected_price = cell_value.value * 0.9
        corrected_price_cell = sheet.cell(row,4)
        corrected_price_cell.value = corrected_price

    wb.save('transactions_test.xlsx')


process_spreedsheet('transactions.xlsx')
