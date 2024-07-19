import pandas as pd
import os

# Create a Pandas Excel writer using openpyxl as the engine.
writer = pd.ExcelWriter('SEEES search request.xlsx', engine='xlsxwriter')

# Iterate over all .tsv files in the current directory
for file in os.listdir('.'):
    if file.endswith('.tsv'):
        try:
            # Read the tsv file into a dataframe
            df = pd.read_csv(file, sep='\t', header=None)
        except pd.errors.ParserError:
            # Handle the anomalous tsv file
            with open(file, 'r') as f:
                lines = f.readlines()
            
            # Write lines to a DataFrame with each line as a single column
            data = {'Content': [line.strip() for line in lines]}
            df = pd.DataFrame(data)
        
        # Use the name of the file (without extension) as the sheet name
        sheet_name = os.path.splitext(file)[0]
        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
# Save the Excel file
writer.close()