import os
import csv
from openpyxl import Workbook

# Create a new workbook (empty)
wb = Workbook()

# Get all files with .tsv extension
tsv_files = [f for f in os.listdir() if f.endswith(".tsv")]

# Flag to track if any data has been written
data_written = False

for filename in tsv_files:
  # Create a new worksheet with filename (without extension)
  sheet = wb.create_sheet(os.path.splitext(filename)[0])
 
  # Open the TSV file in read mode
  with open(filename, 'r') as f:
    # Read data using csv reader with tab delimiter
    reader = csv.reader(f, delimiter='\t')
    # Write data to the worksheet
    for row in reader:
      sheet.append(row)
      data_written = True  # Set flag to indicate data written

# Save the workbook only if data was written
if data_written:
  wb.save("SEEES search request.xlsx")
  print("Successfully combined TSV files into 'SEEES search request.xlsx'")
else:
  print("No data found in TSV files. No Excel file created.")
