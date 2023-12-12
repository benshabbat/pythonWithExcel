import xlwings as xw
import pandas as pd

wk = xw.books.open(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example3\data.xlsx')
sheet = wk.sheets("Sheet")
rg= sheet.range("A1:B1")
print(rg.value)