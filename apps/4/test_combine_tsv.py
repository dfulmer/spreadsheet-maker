import os
import pandas as pd
import pytest
from combine_tsv import combine_tsv_files_to_excel

TEST_DIR = 'test_data'

@pytest.fixture(scope='module')
def setup_tsv_files():
    os.makedirs(TEST_DIR, exist_ok=True)
    test_files = {
        'file1.tsv': 'col1\tcol2\nval1\tval2\n',
        'file2.tsv': 'a\tb\n1\t2\n'
    }
    for filename, content in test_files.items():
        with open(os.path.join(TEST_DIR, filename), 'w') as f:
            f.write(content)

    yield

    for filename in test_files.keys():
        os.remove(os.path.join(TEST_DIR, filename))
    os.rmdir(TEST_DIR)

def test_combine_tsv_files_to_excel(setup_tsv_files):
    os.chdir(TEST_DIR)
    output_filename = 'SEEES search request.xlsx'
    
    combine_tsv_files_to_excel(output_filename)

    assert os.path.exists(output_filename)
    
    xls = pd.ExcelFile(output_filename)
    assert 'file1' in xls.sheet_names
    assert 'file2' in xls.sheet_names
    
    os.remove(output_filename)
    os.chdir('..')
