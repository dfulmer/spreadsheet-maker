import pandas as pd
import glob
import os

# Get a list of all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer
with pd.ExcelWriter('SEEES search request.xlsx') as writer:
    # Loop through each .tsv file
    for file in tsv_files:
        # Read the .tsv file into a Pandas dataframe
        df = pd.read_csv(file, sep='\t')
        # Write the dataframe to a sheet in the Excel file
        df.to_excel(writer, sheet_name=os.path.basename(file[:-4]), index=False)

print("Excel file created successfully!")