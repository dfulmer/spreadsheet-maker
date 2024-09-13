import os
import pandas as pd
import pytest
from tsv_to_excel import combine_tsv_files

@pytest.fixture
def setup_test_files():
    # Create test .tsv files
    pd.DataFrame({'A': [1, 2], 'B': [3, 4]}).to_csv('test1.tsv', sep='\t', index=False)
    pd.DataFrame({'C': [5, 6], 'D': [7, 8]}).to_csv('test2.tsv', sep='\t', index=False)
    yield
    # Clean up test files
    os.remove('test1.tsv')
    os.remove('test2.tsv')
    os.remove('SEEES search request.xlsx')

def test_combine_tsv_files(setup_test_files):
    combine_tsv_files()
    assert os.path.exists('SEEES search request.xlsx')
    
    xl = pd.ExcelFile('SEEES search request.xlsx')
    assert set(xl.sheet_names) == {'test1', 'test2'}
    
    df1 = xl.parse('test1')
    assert df1.to_dict() == {'A': {0: 1, 1: 2}, 'B': {0: 3, 1: 4}}
    
    df2 = xl.parse('test2')
    assert df2.to_dict() == {'C': {0: 5, 1: 6}, 'D': {0: 7, 1: 8}}

def test_no_tsv_files():
    with pytest.raises(ValueError, match="No .tsv files found in the current directory."):
        combine_tsv_files()