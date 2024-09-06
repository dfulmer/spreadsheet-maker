import os
import pandas as pd

def combine_tsv_to_excel(output_filename='SEEES_search_request.xlsx'):
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')
    added_sheets = False
    for filename in os.listdir('.'):
        if filename.endswith('.tsv'):
            df = pd.read_csv(filename, sep='\t')
            sheet_name = os.path.splitext(filename)[0]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            added_sheets = True
    if added_sheets:
        writer.close()
    else:
        # Create an empty workbook if no .tsv files are found
        writer.book.create_sheet(title='Empty')
        writer.close()

if __name__ == "__main__":
    combine_tsv_to_excel()