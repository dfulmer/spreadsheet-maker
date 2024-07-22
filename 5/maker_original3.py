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

    # Create a new Worksheet in the Workbook
    wb.create_sheet(title=filename)

    # Get the new Worksheet object
    ws = wb[filename]

    # Convert DataFrame to a list of lists
    rows = [list(row) for row in dataframe.values]

    # Append the list of lists to the Worksheet
    for row in rows:
        ws.append(row)

# Save and close the Workbook
wb.save('SEEES search request.xlsx')
wb.close()