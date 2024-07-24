import os
import csv
from openpyxl import Workbook

# Create a new workbook
wb = Workbook()
del wb['Sheet']

# Get all files with .tsv extension
tsv_files = [f for f in os.listdir() if f.endswith(".tsv")]

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

# Save the workbook as "SEEES search request.xlsx"
wb.save("SEEES search request.xlsx")

print("Successfully combined TSV files into 'SEEES search request.xlsx'")
