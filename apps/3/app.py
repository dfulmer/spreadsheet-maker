import os
import pandas as pd

def combine_tsv_files():
    # Get a list of all .tsv files in the current directory
    tsv_files = [file for file in os.listdir() if file.endswith('.tsv')]
   
    # Create an Excel writer
    with pd.ExcelWriter('SEEES search request.xlsx') as writer:
        # Iterate over each .tsv file and write it to a separate worksheet
        for file in tsv_files:
            df = pd.read_csv(file, sep='\t')
            df.to_excel(writer, sheet_name=os.path.splitext(file)[0], index=False)

if __name__ == '__main__':
    combine_tsv_files()
