import os
import pandas as pd
import glob

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer using XlsxWriter as the engine
excel_file = 'SEEES search request.xlsx'
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

def read_flexible_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()
    
    # Process lines into a list of dictionaries
    data = []
    for line in lines:
        line = line.strip()
        if ':' in line:
            key, value = line.split(':', 1)
            data.append({'Key': key.strip(), 'Value': value.strip()})
        else:
            data.append({'Key': '', 'Value': line})
    
    return pd.DataFrame(data)

# Loop through all .tsv files
for tsv_file in tsv_files:
    # Read the file using our flexible function
    df = read_flexible_file(tsv_file)
    
    # Get the file name without extension to use as sheet name
    sheet_name = os.path.splitext(tsv_file)[0]
    
    # Write the dataframe to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Close the Pandas Excel writer and save the Excel file
writer.close()

print(f"Excel file '{excel_file}' has been created with {len(tsv_files)} worksheets.")