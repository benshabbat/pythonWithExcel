import pandas as pd

df = pd.read_excel(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example1\data.xlsx')

# Get all Data
results = df.iloc[:]

print(results)