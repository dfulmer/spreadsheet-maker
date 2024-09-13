import os
import pandas as pd
import openpyxl

def combine_tsv_files(output_file='SEEES search request.xlsx'):
    tsv_files = [f for f in os.listdir('.') if f.endswith('.tsv')]
    
    if not tsv_files:
        raise ValueError("No .tsv files found in the current directory.")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for tsv_file in tsv_files:
            df = pd.read_csv(tsv_file, sep='\t')
            sheet_name = os.path.splitext(tsv_file)[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    combine_tsv_files()
    print("Excel file 'SEEES search request.xlsx' has been created.")
