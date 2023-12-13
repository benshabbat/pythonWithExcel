import xlwings as xw
import pandas as pd

wk = xw.Book(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example3\data.xlsx')
sheet = wk.sheets("Sheet")

# rg= sheet.range("A1:B3")
# print(rg.value)

#Get data from dataframe
df= sheet.range("A1:D3").options(pd.DataFrame).value

#how much lines of data
df=df[:2]

#insert data to table
xw.view(df)

print(df)

wk.close