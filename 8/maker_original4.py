import pandas as pd
import glob

# Create a new Excel writer object
excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Get all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Loop through each .tsv file and add it as a new sheet in the Excel file
for tsv_file in tsv_files:
    try:
        # Try to read the .tsv file into a DataFrame
        df = pd.read_csv(tsv_file, sep='\t')
        # Use the file name (without extension) as the sheet name
        sheet_name = tsv_file.split('.')[0]
        # Write the DataFrame to the Excel file without special formatting for the header
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False, header=False)
    except pd.errors.ParserError:
        # If there's a parsing error, read the file as plain text
        with open(tsv_file, 'r') as file:
            lines = file.readlines()
        # Create a new worksheet
        worksheet = excel_writer.book.add_worksheet(tsv_file.split('.')[0])
        # Write each line to a new row in the worksheet
        for row_num, line in enumerate(lines):
            worksheet.write(row_num, 0, line.strip())

# Save the Excel file
excel_writer.close()
