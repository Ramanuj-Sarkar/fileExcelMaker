# Running instructions:
# Install the library openpyxl: pip install openpyxl

# Run this from the command line like:
# python  path\to\cmd_filemaker.py  path\to\input_folder  path\to\excel_file_name

# It looks at all the files in the input folder.
# It deletes the information in the old excel file to create a new one.

import os
import sys
import openpyxl

if __name__ == '__main__':
    print("Why is openpyxl so shy?")