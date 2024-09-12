import os
import pandas as pd
from openpyxl import Workbook

def tsv_to_excel(tsv_files, output_filename):
    """
    Combine multiple .tsv files into a single Excel spreadsheet.

    Arguments:
    - tsv_files (list): A list of paths to the .tsv files.
    - output_filename (str): The name of the output Excel file.

    Returns:
    - None: Creates the spreadsheet and saves it to the provided filename.
    """
    wb = Workbook()
    wb.save(output_filename)

    for index, filename in enumerate(tsv_files):
        dataframe = pd.read_csv(filename, sep='\t', header=None, on_bad_lines='skip')

        filename = filename.rsplit('/', 1)[-1][:-4]

        wb.create_sheet(title=filename)
        ws = wb[filename]
        rows = [list(row) for row in dataframe.values]
        for row in rows:
            ws.append(row)

    wb.save(output_filename)
    wb.close()

if __name__ == '__main__':
    tsv_files = [os.path.join(os.getcwd(), filename) for filename in os.listdir() if filename.endswith('.tsv')]
    if not tsv_files:
        print("No .tsv files found in the current directory.")
        exit(1)
    tsv_to_excel(tsv_files, 'SEEES search request.xlsx')
