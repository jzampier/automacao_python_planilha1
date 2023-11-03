from icecream import ic
import os
import pandas as pd

# path and list of files
path = "bases"
files = os.listdir(path)
# ic(files)

# Create consolidated worksheet
consolidated_worksheet = pd.DataFrame()

# Loop through files
for file_name in files:
    full_path = os.path.join(path, file_name)
    sales_table = pd.read_csv(full_path)
    # ic(sales_table)
    consolidated_worksheet = pd.concat([consolidated_worksheet, sales_table])

consolidated_worksheet = consolidated_worksheet.sort_values(by='first_name')
consolidated_worksheet = consolidated_worksheet.reset_index(drop=True)
ic(consolidated_worksheet)

# Save consolidated worksheet in an Excel file
consolidated_worksheet.to_excel('Sales.xlsx', index=False)

# TODO: Create a routine to send an email with the consolidated sales (Sales.xlsx)
