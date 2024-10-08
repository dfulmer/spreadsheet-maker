# spreadsheet-maker
Makes a spreadsheet

## Getting started

Clone the repository  
```git clone [code from above]``` 

cd into spreadsheet-maker  
```cd spreadsheet-maker```

# Create the virtual environment
```python3 -m venv venv```

# Activate the virtual environment
```source venv/bin/activate```

# Install the packages
```python -m pip install [package name]```  

# Packages needed
pandas  
openpyxl  
xlsxwriter  
pytest  
pytest-cov  


Or install from a requirements.txt file, if there is one:  
```python -m pip install -r requirements.txt```

# Try the program
To use the first program in directory 1:  
This program requires two packages: pandas and xlsxwriter.  
```python -m pip install pandas xlsxwriter```  

Change into the /1 directory:  
```cd 1```  

To run the program you need to put files with .tsv file extensions into the /1 directory. Then give the command:  
```python maker1.py```  

# Spreadsheet Maker Script

# Directory 1
Notes on /1:  
maker_original... are original scripts, unedited.  
maker1.py and maker2.py are edited scripts.  
maker2.py will successfully combine tsv files into one Excel spreadsheet.

maker_original.py - this script does not work. It cannot handle the anomalous tsv file with two lines without tabs.

maker_original2.py - this script does not work either. This one just includes the anomalous tsv file in the final spreadsheet and reports errors processing the other 5 tsv files.

maker_original3.py - this script basically works. All the information from the tsv files are in the final spreadsheet. However, the first row has formatting (bold and centered cells).

maker_original4.py - this script was an attempt to remove the formatting from the first row. While it does remove that formatting, it also has the same first row in every worksheet, including the anomalous worksheet, where the first row of information is missing and replaced with the first row from the other worksheets.

maker1.py - this script has a problem with the anomalous tsv file: the first two lines are missing from the spreadsheet.

maker2.py - this script works. It has all the information from the original tsv files in the spreadsheet. There is no header formatting. The anomalous file is all in a worksheet, including the first two lines. And the next 6 lines of that worksheet have two columns with the information separated out.

# Directory 2
I gave up on this and never got it to work. None of the scripts in this directory create an openable Excel file.

maker_original.py - this script has a problem with the summary tsv file.

maker_original2.py - this script has a problem with the summary tsv file.

maker_original3.py - this script has a problem with the summary tsv file.

maker_original4.py - this script has a problem with the summary tsv file.

maker_original5.py - this script has a problem with the summary tsv file.

# Directory 3
maker_original.py - this script has a problem with the summary tsv file. Although it actually did successfully create a spreadsheet with two worksheets that look okay. It appears to have been in the process of creating the spreadsheet when it got stuck on the summary tsv file which stopped it from proceeding with the creation of all the worksheets.

maker_original2.py - this script succeeded in making the spreadsheet but had a problem with the summary tsv file. It included lines 1 and 2 in the spreadsheet but skipped lines 3-8.

maker_original3.py - this script came close to success but was stripping out leading tabs in many files. This resulted in data getting moved over to the left.

maker1.py - this appears to work. It leaves in those leading tabs and does not format the header, and handles the summary tsv file appropriatley, including all lines in the final spreadsheet. This script is maker_original3.py with one change, line 18. This:
```fields = line.strip().split('\t')```  
was changed to this:
```fields = line.strip('\n').split('\t')```  


# Directory 4
combine_tsv_to_excel.py - has some errors. This script has a problem with the summary tsv file and uses excel_writer.save() instead of excel_writer.close(). I changed engine='openpyxl' to engine='xlsxwriter' too. A spreadsheet is created but it has no data in it.

combine_tsv_to_excel2.py - I think this achieves the desired results. I changed the engine again and also changed writer.save to writer.close. In the worksheet made from the summary file it doesn’t have columns for lines 3 to 8, all the data is in the first cell. The first row does not have any formatting.


# Directory 5
maker_original.py - this script has an error and just creates an empty Excel spreadsheet.

maker_original2.py - this script has an error and just creates an empty Excel spreadsheet.

maker_original3.py - this script has an error and just creates an empty Excel spreadsheet.

maker_original4.py - this script has an error and just creates an empty Excel spreadsheet.

maker_original5.py - this script has a warning in its output. There's a message when you open the spreadsheet created by it. It has a blank worksheet. The summary tsv worksheet only has the first two lines.

maker_original6.py - this script has a lot of the same problems as the last one. There is a warning in its output, a message when you open the spreadsheet crdated by it, it has a blank worksheet, and the summary tsv worksheet only has the first two lines. It does not produce a usable spreadsheet.


# Directory 6
maker_original.py - this is the first try. Doesn't save and doesn't handle the summary file. It creates an Excel file but that's empty.

maker_original2.py - this is the second try. Doesn't save and includes only partial data for the summary file.

maker_original3.py - same as the last one, I just changed .save() to .close(). It does create a spreadsheet successfully, but includes only partial data for the summary file. The header rows are also formatted.

maker_original4.py - This is bad. Now every sheet has Key and Value as column headers.

maker_original5.py - This is really good except the header lines all have formatting and in the summary sheet it starts with “Content” and then the data is below that.

maker_original6.py - This is almost perfect. The summary sheet looks great and each line is just a line. All the sheets have a formatted first row except for the summary sheet, which has no formatting.

maker_original7.py - This does what it should do. All the tsv files are turned into worksheets, the anomalous summary sheet has no columns but does have all data in it, and headers do not have formatting.


# Directory 7
maker_original.py - this is the first try. Everything seems to be there including the summary worksheet. The only problem is that the Excel spreadsheet starts with a blank worksheet, called "Sheet".

maker_original2.py - this is the second try. The result is the same as the first script. I think it might prevent making a sheet if there was a blank tsv file, but that isn't the problem. Actually no, it just doesn't create a spreadsheet if all the tsv files are empty.

maker1.py - this is the first try plus one line I added to delete the first blank sheet. This appears to work as expected. All the worksheets are there and the first row does not have formatting.


# Directory 8
maker_original.py - this had a problem with the summary file and the two first lines that don't have tab characters. The script generates an error and while a spreadsheet is created, it is empyt and has an error when opened. It never worked.

maker_original2.py - this one creates a spreadsheet with all the worksheets. The first row of most worksheets is formatted. For the summary tsv file, everything is put into one cell. It has carriage returns and everything but it is all in one cell that appears as one long string in the spreadsheet.

maker_original3.py - this is almost perfect, it just has the header formatted in bold (I also changed .save to .close. The error was "AttributeError: 'XlsxWriter' object has no attribute 'save'").

maker_original4.py - this one does not have a first row, it removed it. It was the row with the header information.

maker_original5.py - this one does what I want. The summary tsv file is all in a worksheet, and the header is not formatted.


# Spreadsheet Maker App

# Directory apps/1
combine_tsv_to_excel.py - this didn’t really work at first cause it’s just a function.

combine_tsv_to_excel2.py - this did create a spreadsheet but the first row is formatted. The Summary has a ‘0’ in the second column of the 2nd row for some reason. There should be three equals signs. There is an error when you open the Excel spreadsheet.

pytest - this command did run the test. I had to edit the original file apps/1/test_combine_tsv_to_excel.py because I changed the name of the main file. So I added this:  
from combine_tsv_to_excel2 import combine_tsv_to_excel  
(note the “2”)  
There’s only one test, and it passed.

pytest --cov=combine_tsv_to_excel2 - this runs the test and gives a little coverage report in the output.

pytest --cov=combine_tsv_to_excel2 --cov-report=html - this runs the test and puts an html report in htmlcov directory.

# Directory apps/2
combine_tsv.py - had to change .save to .close but then it worked. I also updated my test files with this one. There were problems caused by using equals signs - the spreadsheet showed an error on opening. I think that may have explained all the previous problems where there was an error when you open the Excel spreadsheet. This script now makes a spreadsheet, but it does have formatting of the first row - bold text, and centered in the cell.

test_combine_tsv.py - this has two tests.

pytest - one test passes, the second test does not. There seems to be other errors as well.

pytest --cov=combine_tsv --cov-report=html - This creates a report.

combine_tsv2.py - I made a second attempt.

test_combine_tsv2.py - this has two tests. It still has one that pass and one that does not.

# Directory apps/3
app.py - this works almost perfectly, it just has a formatted first row.

pytest test_app.py - both tests fail.

pytest test_app2.py - both tests fail. I just changed the first script by adding an import statement to get the function being tested.

pytest test_app3.py - both tests pass. But this is working in the same directory and creating a bunch of test files there. I added some commands to remove the test files. Also, the second test which tests that when there are no tsv files, no xlsx file is created, isn't really working. I had to put in a statement to remove a blank file because the combine_tsv_files function actually does create a spreadsheet even if there are not tsv files. I took out the tmpdir stuff, which I don't quite understand yet.

pytest --cov=app test_app3.py - this command prints out a little report. 90%.

pytest --cov=app test_app3.py --cov-report=html - this creates the html report.

Hint: to get to 100%, do this: in app.py, on line 16, after combine_tsv_files() ... Add this: "# pragma: no cover"

# Directory apps/4
python combine_tsv.py - this works almost perfectly, it just has a formatted first row.

pytest test_combine_tsv.py - this runs one tests, which passes. The test file is creating a subdirectory and two files which it deletes at the end of the test. 

pytest --cov=combine_tsv --cov-report=html = this generates the coverage report.


# Directory apps/5
python tsv_to_excel.py - after making some edits, this worked. The spreadsheet has one blank sheet to begin with but now first row formatting.

Tests - I had a tests file but it didn't work at all, one test FAILED, one had ERROR.


# Directory apps/6
python tsv_to_excel.py - this creates a spreadsheet. It also has an error if there are no tsv files. The spreadsheet is complete but does have formatting of the first row.

pytest test_tsv_to_excel.py - this seems not to work if there are any .tsv files in the same directory. If there are no .tsv files in the same directory then the two tests pass.

pytest --cov=tsv_to_excel test_tsv_to_excel.py - this produces a coverage report showing 87% coverage.

pytest --cov=tsv_to_excel --cov-report=html test_tsv_to_excel.py - this generates the html report.


# Directory apps/7
app.py - this did create a spreadsheet. But there was an error on opening. There is a blank worksheet. The second cell of the second row of the summary tsv isn’t in the spreadsheet. Otherwise the data seems to be there.

pytest test_app.py - this runs the test, which fails.

pytest --cov=app test_app.py - this prints out a little coverage report. 92%


# Directory apps/8
python combine_tsv_to_excel.py - this creates a spreadsheet. It also has an sensible error message if there are no tsv files. The spreadsheet is complete but does have formatting of the first row.

pytest test_combine_tsv_to_excel.py - originally this didn't really work. I made some changes so that there are now 2 generic tests to show formatting, then 3 tests which pass. It used 'tmpdir' and I switched it over to tmp_path_factory.

pytest --cov=combine_tsv_to_excel --cov-report html = generates the coverage report.


# Deactivate the virtual environment
```deactivate```
