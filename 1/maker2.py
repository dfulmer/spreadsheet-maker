# This is maker_original3.py plus the header formatting removal using pandas.io.formats.excel

import os
import pandas as pd
import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style=None

def read_tsv_file(filename):
    try:
        # Attempt to read the TSV file normally
        df = pd.read_csv(filename, sep='\t')
    except pd.errors.ParserError:
        # If there is a parsing error, manually process the file
        with open(filename, 'r') as file:
            lines = file.readlines()
        
        data = []
        for line in lines:
            data.append(line.strip().split('\t'))
        
        # Determine the number of columns from the longest row
        max_columns = max(len(row) for row in data)
        
        # Ensure all rows have the same number of columns
        for row in data:
            while len(row) < max_columns:
                row.append('')
        
        df = pd.DataFrame(data[1:], columns=data[0])
    
    return df

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Loop through all files in the current directory
for filename in os.listdir('.'):
    if filename.endswith('.tsv'):
        # Use the custom read_tsv_file function
        df = read_tsv_file(filename)
        # Use the filename (without extension) as the sheet name
        sheet_name = os.path.splitext(filename)[0]
        # Write the DataFrame to a new sheet in the Excel file
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the Excel file
writer.close()
