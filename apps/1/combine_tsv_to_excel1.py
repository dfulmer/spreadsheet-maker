import os
import pandas as pd

def combine_tsv_to_excel(output_filename="SEEES search request.xlsx"):
    # List all .tsv files in the current directory
    tsv_files = [f for f in os.listdir() if f.endswith('.tsv')]
    
    if not tsv_files:
        raise FileNotFoundError("No TSV files found in the current directory.")
    
    # Create a Pandas Excel writer
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        # Loop over all TSV files and add each as a separate sheet in the Excel workbook
        for tsv_file in tsv_files:
            # Read the TSV file
            df = pd.read_csv(tsv_file, sep='\t')
            # Write the DataFrame to a specific sheet named after the file (without extension)
            sheet_name = os.path.splitext(tsv_file)[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    return output_filename
