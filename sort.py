import pandas as pd
from openpyxl import load_workbook

# TODO: improve for dynamic names and sheet names
excel_file = 'Documents/report.xlsv'
sheet_name = 'sheet'

df = pd.read_excel(excel_file, sheet_name=sheet_name)

# define columns
def custom_sort(row):
  return (row['Order No'], row['Order Line'], row['Date Entered'])

# sort them
df = df.sort_values(by=[custom_sort])

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
  writer.book = load_workbook(excel_file)
  df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Custom sorting is complete.")
