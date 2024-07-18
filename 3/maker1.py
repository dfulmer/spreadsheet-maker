import pandas as pd
import glob
import os

# Get a list of all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer
with pd.ExcelWriter('SEEES search request.xlsx') as writer:
    # Loop through each .tsv file
    for file in tsv_files:
        # Initialize an empty list to store rows
        rows = []
        # Open the file and read it line by line
        with open(file, 'r') as f:
            for line in f:
                # Split the line into fields (assuming tab-separated values)
                fields = line.strip('\n').split('\t')
                # If the line has no tabs, use the entire line as a single field
                if len(fields) == 1:
                    fields = [line.strip()]
                # Add the fields to the rows list
                rows.append(fields)
        # Create a dataframe from the rows list
        df = pd.DataFrame(rows)
        # Write the dataframe to a sheet in the Excel file
        df.to_excel(writer, sheet_name=os.path.basename(file[:-4]), index=False, header=False)

print("Excel file created successfully!")