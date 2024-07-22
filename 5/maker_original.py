import glob
import pandas as pd
from openpyxl import Workbook

# Create a new Excel Workbook
wb = Workbook()

# Save it with the name 'SEEES search request'
wb.save('SEEES search request.xlsx')

# Get a list of all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Iterate through each tsv file and convert it to an Excel worksheet
for index, filename in enumerate(tsv_files):
    # Read the tsv file into a Pandas DataFrame
    dataframe = pd.read_csv(filename, sep='\t', header=None)
    
    # Add DataFrame to Excel Workbook as a new Worksheet
    wb.create_sheet(dataframe)
    
    # Rename the Worksheet with the name of the original file
    wb['dataframe'] = filename
    
# Save and close the Workbook
wb.save('SEEES search request.xlsx')
wb.close()
