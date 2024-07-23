import os
import pandas as pd
import glob
import csv

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_file = 'SEEES search request.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

def read_normal_tsv(file_path):
    return pd.read_csv(file_path, sep='\t', encoding='utf-8')

def read_anomalous_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Process lines into a single column DataFrame without a header
    data = [line.strip() for line in lines]
    return pd.DataFrame(data, columns=[''])

# Loop through all .tsv files
for tsv_file in tsv_files:
    try:
        # Try to read the .tsv file normally first
        df = read_normal_tsv(tsv_file)
        is_anomalous = False
    except pd.errors.ParserError:
        # If that fails, use our anomalous file reading function
        df = read_anomalous_file(tsv_file)
        is_anomalous = True
    
    # Get the file name without extension to use as sheet name
    sheet_name = os.path.splitext(tsv_file)[0]
    
    # Write the dataframe to the Excel file
    if is_anomalous:
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    else:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Close the Pandas Excel writer and save the Excel file
writer.close()

print(f"Excel file '{excel_file}' has been created with {len(tsv_files)} worksheets.")