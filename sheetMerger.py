import pandas as pd

# Specify the path to your Excel file
excel_file = 'Vehicle information_20241129_113139.xlsx' #Paste the DMS output file name

# Read all sheets into a dictionary of DataFrames
sheet_data = pd.read_excel(excel_file, sheet_name=None)

# Initialize an empty DataFrame for merging
merged_data = pd.DataFrame()

# Iterate through each sheet and append its data
for sheet_name, data in sheet_data.items():
    data['SheetName'] = sheet_name  # Optional: Add a column to track the sheet name
    merged_data = pd.concat([merged_data, data], ignore_index=True)

# Save the merged data into a new Excel file
merged_file = 'result.xlsx'
merged_data.to_excel(merged_file, index=False)

print(f"All sheets have been merged into {merged_file}")