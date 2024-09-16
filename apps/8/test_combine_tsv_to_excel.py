import os
import pytest
from combine_tsv_to_excel import combine_tsv_to_excel
import pandas as pd

# The first two are just generic samples to show the tmp_path format
def test_create_file(tmp_path):
    d = tmp_path / "sub"
    d.mkdir()
    p = d / "hello.txt"
    p.write_text("CONTENT", encoding="utf-8")
    assert p.read_text(encoding="utf-8") == "CONTENT"
    assert len(list(tmp_path.iterdir())) == 1
    #assert 0

def test_create_file2(tmp_path):
    #d = tmp_path / "sub"
    #d.mkdir()
    q = tmp_path / "hello2.txt"
    q.write_text("CONTENT", encoding="utf-8")
    assert q.read_text(encoding="utf-8") == "CONTENT"
    #assert len(list(tmp_path.iterdir())) == 1


@pytest.fixture
def create_dummy_tsv_files(tmp_path_factory):
    path = tmp_path_factory.mktemp("sub")
    # Create dummy .tsv files for testing
    #tsv1 = tmp_path.join("file1.tsv")
    tsv1 = path / "file1.tsv"
    tsv1.write_text("Column1\tColumn2\nValue1\tValue2\n")

    tsv2 = path / "file2.tsv"
    tsv2.write_text("Column1\tColumn2\nValue3\tValue4\n")

    return tsv1, tsv2, path

@pytest.fixture
def create_totally_empty_folder(tmp_path_factory):
    pathb = tmp_path_factory.mktemp("subb")
    return pathb

def test_combine_tsv_to_excel(create_dummy_tsv_files):
    tsv1, tsv2, patha = create_dummy_tsv_files
    os.chdir(patha)  # Change to the temporary directory

    combine_tsv_to_excel()

    assert os.path.exists("SEEES_search_request.xlsx")

def test_no_tsv_files(create_totally_empty_folder):
    path = create_totally_empty_folder
    os.chdir(path)
    combine_tsv_to_excel()
    assert os.path.exists("SEEES_search_request.xlsx") == False

def test_sheet_names(create_dummy_tsv_files):
    tsv1, tsv2, patha = create_dummy_tsv_files
    os.chdir(patha)

    combine_tsv_to_excel()

    excel_file = pd.ExcelFile("SEEES_search_request.xlsx")
    sheet_names = excel_file.sheet_names
    assert sheet_names == ["file1", "file2"]
