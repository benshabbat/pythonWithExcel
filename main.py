from openpyxl.workbook import Workbook
from openpyxl import load_workbook



# Create a new workbook object
wb= Workbook()


# load existing spreadsheet
wb=load_workbook('data.xlsx')


# Create a active worksheet
ws = wb.active


# Print something from our Spreadsheet
print(ws["A2"].value)