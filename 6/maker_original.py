import os
import pandas as pd
import glob

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_file = 'SEEES search request.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# Loop through all .tsv files
for tsv_file in tsv_files:
    # Read the .tsv file
    df = pd.read_csv(tsv_file, sep='\t')
    
    # Get the file name without extension to use as sheet name
    sheet_name = os.path.splitext(tsv_file)[0]
    
    # Write the dataframe to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer.save()

print(f"Excel file '{excel_file}' has been created with {len(tsv_files)} worksheets.")