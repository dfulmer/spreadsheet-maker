import os
import pandas as pd

def combine_tsv_to_excel():
    # Get a list of all .tsv files in the current directory
    tsv_files = [filename for filename in os.listdir() if filename.endswith('.tsv')]

    if not tsv_files:
        print("No .tsv files found in the current directory.")
        return

    # Create a Pandas Excel writer
    excel_writer = pd.ExcelWriter('SEEES_search_request.xlsx', engine='xlsxwriter')

    # Iterate through each .tsv file and add it as a separate worksheet
    for tsv_file in tsv_files:
        sheet_name = os.path.splitext(tsv_file)[0]  # Use the filename without extension as the sheet name
        df = pd.read_csv(tsv_file, sep='\t')
        df.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    # Save the Excel file
    excel_writer.close()
    print(f"Combined {len(tsv_files)} .tsv files into SEEES_search_request.xlsx")

if __name__ == "__main__":
    combine_tsv_to_excel()
