import os
import pandas as pd

# Get a list of all .tsv files in the current directory
tsv_files = [file for file in os.listdir() if file.endswith('.tsv')]

# Create a Pandas Excel writer
excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Write each .tsv file to a separate worksheet
for tsv_file in tsv_files:
    sheet_name = os.path.splitext(tsv_file)[0]  # Use the filename without extension as the sheet name
    df = pd.read_csv(tsv_file, sep='\t', header=None)  # Read the entire file as a single cell
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Save the Excel file
excel_writer.save()

print(f"Excel spreadsheet 'SEEES_search_request.xlsx' created with {len(tsv_files)} worksheets.")
