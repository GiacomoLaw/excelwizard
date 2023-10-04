import pandas as pd

# TODO: dynamic input
input_csv = 'C:/Users/marketing/Documents/report.csv'
output_csv = 'C:/Users/marketing/Documents/output.csv'

# add csv to datafram
df = pd.read_csv(input_csv, encoding='cp1252') # encoding for windows

# define sorting columns
df.sort_values(by=['Order No', 'Order Line', 'Date Entered'], inplace=True)

# save to output csv
df.to_csv(output_csv, index=False)

print(f'Sorted and saved to {output_csv}')
