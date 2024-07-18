import pandas as pd
import glob
import os
import io

# Get a list of all .tsv files in the current directory
tsv_files = glob.glob('*.tsv')

# Create a Pandas Excel writer
with pd.ExcelWriter('SEEES search request.xlsx') as writer:
    # Loop through each .tsv file
    for file in tsv_files:
        # Read the .tsv file into a Pandas dataframe
        with open(file, 'r') as f:
            content = f.read()
            # Replace any occurrences of two consecutive newlines with a tab character
            content = content.replace('\n\n', '\n\t')
            # Use io.StringIO to create a file-like object
            df = pd.read_csv(io.StringIO(content), sep='\t', on_bad_lines='warn')
        # Write the dataframe to a sheet in the Excel file
        df.to_excel(writer, sheet_name=os.path.basename(file[:-4]), index=False)

print("Excel file created successfully!")