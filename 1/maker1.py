import os
import pandas as pd
import fnmatch
import pandas.io.formats.excel
pandas.io.formats.excel.ExcelFormatter.header_style=None

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Loop through all files in the current directory
for filename in os.listdir('.'):
    if filename.endswith('.tsv'):
        # Read each .tsv file into a DataFrame

        # This works but unfortunately cuts off the first 2 lines of the summary tsv file.
        if fnmatch.fnmatch(filename, '*summary*'):
            df = pd.read_csv(filename, sep='\t', header=2)
        else:
            df = pd.read_csv(filename, sep='\t')
        # Use the filename (without extension) as the sheet name
        sheet_name = os.path.splitext(filename)[0]
        # Write the DataFrame to a new sheet in the Excel file        
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# Save the Excel file
writer.close()
