import pandas as pd
import glob
import os

# Set the directory where the .tsv files are located (current directory in this case)
directory = '.'

# Collect all .tsv files in the directory
tsv_files = glob.glob(os.path.join(directory, '*.tsv'))

# Create a writer object to write the combined Excel file
#excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='openpyxl')
# I'm changing this
excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Loop through the list of .tsv files and add each as a separate sheet in the Excel file
for tsv_file in tsv_files:
    # Get the base name of the file without extension to use as sheet name
    sheet_name = os.path.basename(tsv_file).replace('.tsv', '')
    
    # Read the .tsv file into a DataFrame
    df = pd.read_csv(tsv_file, delimiter='\t')
    
    # Write the DataFrame to a sheet in the Excel file
    df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

# Save the Excel file
excel_writer.save()