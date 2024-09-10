import os
import pandas as pd

def combine_tsv_files_to_excel(output_filename='SEEES search request.xlsx'):
    # Get list of .tsv files in the current directory
    tsv_files = [f for f in os.listdir('.') if f.endswith('.tsv')]
    
    # Create a new Excel writer object
    with pd.ExcelWriter(output_filename) as writer:
        for tsv_file in tsv_files:
            df = pd.read_csv(tsv_file, delimiter='\t')
            sheet_name = os.path.splitext(tsv_file)[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    combine_tsv_files_to_excel()
