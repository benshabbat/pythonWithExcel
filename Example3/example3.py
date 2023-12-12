import xlwings as xw
# import pandas as pd

wk = xw.Book(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example3\data.xlsx')
sheet = wk.sheets("Sheet")
rg= sheet.range("A1:B3")
print(rg.value)