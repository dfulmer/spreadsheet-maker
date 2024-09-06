import os
import pandas as pd
import pytest
from combine_tsv2 import combine_tsv_to_excel

@pytest.fixture
def setup_tsv_files(tmpdir):
    filenames = ['file1.tsv', 'file2.tsv']
    for filename in filenames:
        df = pd.DataFrame({'col1': [1, 2], 'col2': [3, 4]})
        df.to_csv(tmpdir.join(filename), sep='\t', index=False)
    return tmpdir

def test_combine_tsv_to_excel(setup_tsv_files):
    output_filename = setup_tsv_files.join('SEEES_search_request.xlsx')
    os.chdir(setup_tsv_files)
    combine_tsv_to_excel(output_filename)
    assert os.path.exists(output_filename)
    xls = pd.ExcelFile(output_filename)
    assert 'file1' in xls.sheet_names
    assert 'file2' in xls.sheet_names

def test_empty_directory(tmpdir):
    output_filename = tmpdir.join('SEEES_search_request.xlsx')
    os.chdir(tmpdir)
    combine_tsv_to_excel(output_filename)
    assert os.path.exists(output_filename)
    xls = pd.ExcelFile(output_filename)
    assert len(xls.sheet_names) == 0
