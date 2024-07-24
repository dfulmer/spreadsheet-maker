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

Notes on /1:
maker_original... are original scripts, unedited.
maker1.py and maker2.py are edited scripts.
maker2.py will successfully combine tsv files into one Excel spreadsheet.

# Directory 2
I gave up on this and never got it to work.

# Directory 3
maker_original3.py came close but was stripping out leading tabs in many files.
maker1.py - this appears to work. It leaves in those leading tabs and does not format the header, and handles the summary tsv file appropriatley, including all lines in the final spreadsheet.

# Directory 4
combine_tsv_to_excel.py - has some errors
combine_tsv_to_excel2.py - I think this achieves the desired results.

# Directory 5

# Directory 6
maker_original.py - this is the first try. Doesn't save and doesn't handle the summary file.

maker_original2.py - this is the second try. Doesn't save and includes only partial data for the summary file.

maker_original3.py - same as the last one, just changes .save() to .close()

maker_original4.py - This is bad. Now every sheet has Key and Value as column headers.

maker_original5.py - This is really good except the header lines all have formatting and in the summary sheet it starts with “Content” and then the data is below that.

maker_original6.py - This is almost perfect. The summary sheet looks great and each line is just a line. 

maker_original7.py - This does what it should do. All the tsv files are turned into worksheets, the anomalous summary sheet has no columns but does have all data in it, and headers do not have formatting.


# Directory 7
maker_original.py - this is the first try. Everything seems to be there except the Excel spreadsheet starts with a blank worksheet, called "Sheet".

maker_original2.py - this is the second try. I think it would prevent making a sheet if there was a blank tsv file, but that isn't the problem. Actually no, it just doesn't create a spreadsheet if all the tsv files are empty.

maker1.py - 

TBD

# Deactivate the virtual environment
```deactivate```
