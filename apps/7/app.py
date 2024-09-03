import openpyxl
import glob

def combine_tsv_files_to_excel(output_filename="SEEES search request.xlsx"):
    """Combines all .tsv files in the current directory into a single Excel spreadsheet.

    Args:
        output_filename (str, optional): The name of the output Excel file. Defaults to "SEEES search request.xlsx".
    """

    workbook = openpyxl.Workbook()
    for filename in glob.glob("*.tsv"):
        worksheet = workbook.create_sheet(title=filename.split(".")[0])
        with open(filename, "r") as tsv_file:
            for row_index, row in enumerate(tsv_file, start=1):
                worksheet.append(row.strip().split("\t"))
    workbook.save(output_filename)

if __name__ == "__main__":
    combine_tsv_files_to_excel()
