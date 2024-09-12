import os
import shutil
import tempfile
import pytest

from tsv_to_excel import tsv_to_excel

def setup_module():
    # Create a temporary directory for tests
temp_dir = tempfile.mkdtemp()
    os.makedirs(os.path.join(temp_dir, 'input'), exist_ok=True)
    os.makedirs(os.path.join(temp_dir, 'output'), exist_ok=True)

    # Create test files
file_names = ['test1.tsv', 'test2.tsv']
    for filename in file_names:
        with open(os.path.join(temp_dir, 'input', filename), 'w') as file:
            file.write('a b c\n1 2 3\n4 5 6\n')
    return temp_dir

def teardown_module(temp_dir):
    # Remove the temporary directory after tests are completed
shutil.rmtree(temp_dir)

def test_tsv_to_excel():
    # Get the paths of the test files
tsv_files = [os.path.join(temp_dir, 'input', filename) for filename in os.listdir(os.path.join(temp_dir, 'input'))]

    # Call the function to create the Excel spreadsheet
tsv_to_excel(tsv_files, os.path.join(temp_dir, 'output', 'SEEES search request.xlsx'))

    # Assert that the output file exists
assert os.path.exists(os.path.join(temp_dir, 'output', 'SEEES search request.xlsx'))

    # Open the output file and check the worksheet names
wb = openpyxl.load_workbook(os.path.join(temp_dir, 'output', 'SEEES search request.xlsx'))
    assert wb.sheetnames == ['test1', 'test2']
