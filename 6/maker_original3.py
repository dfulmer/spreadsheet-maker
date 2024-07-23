import os
import pandas as pd
import glob
import csv

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_file = 'SEEES search request.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# Function to read TSV files more flexibly
def read_flexible_tsv(file_path):
    with open(file_path, 'r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        data = list(reader)
    
    # Find the first row with tabs (assuming it's the header)
    header_index = next((i for i, row in enumerate(data) if '\t' in ''.join(row)), 0)
    
    # If we found a header with tabs, split all subsequent rows by tabs
    if header_index > 0:
        header = data[header_index][0].split('\t')
        rows = [row[0].split('\t') for row in data[header_index+1:]]
        return pd.DataFrame(rows, columns=header)
    else:
        # If no header with tabs found, treat the whole file as a single column
        return pd.DataFrame(data, columns=['Content'])

# Loop through all .tsv files
for tsv_file in tsv_files:
    try:
        # Try to read the .tsv file normally first
        df = pd.read_csv(tsv_file, sep='\t', encoding='utf-8')
    except pd.errors.ParserError:
        # If that fails, use our flexible reading function
        df = read_flexible_tsv(tsv_file)
    
    # Get the file name without extension to use as sheet name
    sheet_name = os.path.splitext(tsv_file)[0]
    
    # Write the dataframe to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Close the Pandas Excel writer and save the Excel file
writer.close()

print(f"Excel file '{excel_file}' has been created with {len(tsv_files)} worksheets.")