import pytest
import pandas as pd
import os

def test_combine_tsv_files(tmpdir):
    # Create some sample .tsv files
    tsv_files = ['file1.tsv', 'file2.tsv']
    for file in tsv_files:
        with open(os.path.join(tmpdir, file), 'w') as f:
            f.write('col1\tcol2\n')
            f.write('1\t2\n')
   
    # Run the function
    combine_tsv_files()
   
    # Check that the Excel file was created
    assert os.path.exists('SEEES search request.xlsx')
   
    # Check that each .tsv file has a corresponding worksheet
    with pd.ExcelFile('SEEES search request.xlsx') as xlsx:
        assert len(xlsx.sheet_names) == len(tsv_files)
        for file in tsv_files:
            sheet_name = os.path.splitext(file)[0]
            assert sheet_name in xlsx.sheet_names

def test_empty_directory():
    # Run the function in an empty directory
    combine_tsv_files()
   
    # Check that no Excel file was created
    assert not os.path.exists('SEEES search request.xlsx')