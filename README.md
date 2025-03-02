A Python script that converts a GEDCOM file (a standard format for genealogical data) to an Excel file using the gedcom library for parsing and pandas for creating the Excel output

A pre-processing stage will attempt to repair errors in the input GEDCOM file such as level number jumping and invalid characters

Before running this code, you need to install the required libraries:
pip install python-gedcom pandas openpyxl

Usage: ged2excel.py [-h] input_file output_file

Convert a GEDCOM file to an Excel file.

positional arguments:
  input_file    Path to the input GEDCOM file (e.g., family.ged)
  output_file   Path to the output Excel file (e.g., output.xlsx)

optional arguments:
  -h, --help    show this help message and exit
