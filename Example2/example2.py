import pandas as pd

df = pd.read_excel(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example1\data.xlsx')

results = df.columns

print(results)