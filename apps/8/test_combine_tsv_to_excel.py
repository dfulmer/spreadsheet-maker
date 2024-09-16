import os
import pytest
from combine_tsv_to_excel import combine_tsv_to_excel

@pytest.fixture
def create_dummy_tsv_files(tmpdir):
    # Create dummy .tsv files for testing
    tsv1 = tmpdir.join("file1.tsv")
    tsv1.write("Column1\tColumn2\nValue1\tValue2\n")

    tsv2 = tmpdir.join("file2.tsv")
    tsv2.write("Column1\tColumn2\nValue3\tValue4\n")

    return tsv1, tsv2

def test_combine_tsv_to_excel(create_dummy_tsv_files):
    tsv1, tsv2 = create_dummy_tsv_files
    os.chdir(tmpdir)  # Change to the temporary directory

    combine_tsv_to_excel()

    assert os.path.exists("SEEES_search_request.xlsx")

def test_no_tsv_files():
    with pytest.raises(SystemExit):
        combine_tsv_to_excel()

def test_sheet_names(create_dummy_tsv_files):
    tsv1, tsv2 = create_dummy_tsv_files
    os.chdir(tmpdir)

    combine_tsv_to_excel()

    excel_file = pd.ExcelFile("SEEES_search_request.xlsx")
    sheet_names = excel_file.sheet_names
    assert sheet_names == ["file1", "file2"]

# Add more tests as needed
