import pandas as pd
import glob

# Create a new Excel writer object
excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Loop through each .tsv file and add it as a new sheet in the Excel file
for tsv_file in tsv_files:
    # Read the .tsv file into a DataFrame
    df = pd.read_csv(tsv_file, sep='\t')
    
    # Use the file name (without extension) as the sheet name
    sheet_name = tsv_file.split('.')[0]
    
    # Write the DataFrame to the Excel file
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Save the Excel file
excel_writer.save()
