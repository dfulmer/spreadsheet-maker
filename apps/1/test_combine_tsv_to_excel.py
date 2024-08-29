import pytest
import os
import pandas as pd
from combine_tsv_to_excel2 import combine_tsv_to_excel

# Fixture to create sample TSV files for testing
@pytest.fixture
def create_sample_tsv_files(tmpdir):
    # Create some sample TSV files in the temporary directory
    file1 = tmpdir.join("file1.tsv")
    file2 = tmpdir.join("file2.tsv")
    
    # Write sample data to these files
    file1.write("col1\tcol2\n1\t2\n3\t4")
    file2.write("col1\tcol2\n5\t6\n7\t8")
    
    # Change the working directory to the temporary one
    os.chdir(tmpdir)
    
    yield
    
    # Clean up after the test (Pytest does this automatically, but we're making sure)
    os.chdir("..")

# Test the combine_tsv_to_excel function
def test_combine_tsv_to_excel(create_sample_tsv_files):
    # Call the function
    output_file = combine_tsv_to_excel()
    
    # Check if the output file was created
    assert os.path.exists(output_file), "Output Excel file was not created."
    
    # Load the Excel file and check the content
    with pd.ExcelFile(output_file) as xls:
        # Check that the sheets correspond to the TSV file names
        assert "file1" in xls.sheet_names, "Sheet for file1.tsv is missing."
        assert "file2" in xls.sheet_names, "Sheet for file2.tsv is missing."
        
        # Check that the data was written correctly
        df1 = pd.read_excel(xls, sheet_name="file1")
        df2 = pd.read_excel(xls, sheet_name="file2")
        
        assert df1.equals(pd.DataFrame({"col1": [1, 3], "col2": [2, 4]})), "Data in file1 sheet is incorrect."
        assert df2.equals(pd.DataFrame({"col1": [5, 7], "col2": [6, 8]})), "Data in file2 sheet is incorrect."
