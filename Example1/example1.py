import openpyxl

book = openpyxl.load_workbook(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example1\data.xlsx')
sheet = book["Sheet"]
sheet["A2"].value ="bb"
# Get all Data
for row in sheet.iter_rows(values_only=True):
    print(row)