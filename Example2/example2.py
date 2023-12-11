import pandas as pd

df = pd.read_excel(r'C:\Users\bensh\OneDrive\שולחן העבודה\pythonWithExcel\Example2\data.xlsx')

# Get all Data
results = df.iloc[:]

# print(results)

# Get by filter
# results = df[df["קוד"].str.match("צהוב")]


# Get by contains
results = df[df["קוד"].str.contains("ו")]

print(results)