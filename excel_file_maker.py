# takes in all the files of a folder
# copies them into different tabs of an excel file
# puts the excel file into the same folder
import os
import sys
import pandas as pd
import openpyxl

for arg in sys.argv:
    print(arg)

print("End of arguments")

directory = './file_input'
new_excel = './file_input/output.xlsx'

workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'] = 'TEMP'
workbook.save(new_excel)

print(os.listdir(directory))

text_files = [x for x in os.listdir(directory) if x[-4:] == '.txt']

print(text_files)

for sheetnum, name in enumerate(text_files, 1):
    lines = []
    # Open file
    with open(os.path.join(directory, name), 'r') as f:
        # Read content of file
        lines += f.read().split('\n')

    excel_df = [x.split(' ') for x in lines]

    workbook = openpyxl.load_workbook(new_excel)
    workbook.create_sheet(f'{name}_{sheetnum}')
    sheet = workbook[f'{name}_{sheetnum}']

    for row in excel_df:
        sheet.append(row)

    # for row in excel_df:
    workbook.save(new_excel)

workbook.remove(workbook['Sheet'])
workbook.save(new_excel)




