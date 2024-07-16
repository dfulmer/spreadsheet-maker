import os
import pandas as pd

def read_tsv_file(filename):
    with open(filename, 'r') as file:
        lines = file.readlines()
    
    # Separate header and data lines
    header = None
    data = []
    
    for line in lines:
        if '\t' in line:
            if header is None:
                header = line.strip().split('\t')
            else:
                data.append(line.strip().split('\t'))
        else:
            # Handle lines without tabs as separate rows in a different manner if needed
            data.append([line.strip()])
    
    return pd.DataFrame(data, columns=header if header else None)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

# Loop through all files in the current directory
for filename in os.listdir('.'):
    if filename.endswith('.tsv'):
        try:
            # Use the custom read_tsv_file function
            df = read_tsv_file(filename)
            # Use the filename (without extension) as the sheet name
            sheet_name = os.path.splitext(filename)[0]
            # Write the DataFrame to a new sheet in the Excel file
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"Error processing file {filename}: {e}")

# Save the Excel file
writer.close()
