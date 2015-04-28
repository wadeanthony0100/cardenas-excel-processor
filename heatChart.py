"""
title: heatChart.py
author: Wade Mauger
date: 4/22/2015
description:Utilizes the openpyxl library
to open and process an Excel file and
output an appropriate Excel visual file.

"""
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.styles import fills, Color, Fill, Style, PatternFill
from openpyxl.styles.colors import GREEN, BLACK

#define how a filled cell is to be filled
from openpyxl.utils import get_column_letter

myFill = PatternFill(patternType=fills.FILL_SOLID, fgColor=GREEN)
backFill = PatternFill(patternType=fills.FILL_SOLID, fgColor=BLACK)

#collect I/O data
input_name = input("Input file name- This must be placed \nin the directory above this script: ")
output_name = input("Output file name: ")
vertical_letter = input("Which column of the input file is meant to be the horizontal axis? ").upper()
horizontal_letter = input("Which column of the input file is meant to be the vertical axis? ").upper()

#ensure proper formatting
if input_name[-5:] != '.xlsx':
    input_name += '.xlsx'

if output_name[-5:] != '.xlsx':
    output_name += '.xlsx'

#initialize workbook, get data worksheet
wb = load_workbook(filename = input_name,  use_iterators=True)
data_sheet = wb.get_sheet_by_name('Data')

#get the number of valid rows in the sheet, for iteration
num_rows = data_sheet.max_row

#initialize a dictionary of meaning strings
#data will be stored as a mapping of meaning string keys
#to lists of arrows for which the meaning exists
meanings_to_lst_arrow = {}

max_arrow = 0

for x in range(2, num_rows+1):
    arrow = data_sheet[vertical_letter + str(x)].value
    meaning = data_sheet[horizontal_letter + str(x)].value
    if arrow > max_arrow:
        max_arrow = arrow
    if meaning in meanings_to_lst_arrow:
        if not arrow in meanings_to_lst_arrow[meaning]:
            meanings_to_lst_arrow[meaning].append(arrow)
    else:
        meanings_to_lst_arrow[meaning] = [arrow]

#print(meanings_to_lst_arrow)

#make a new Worksheet
new_book = Workbook()
new_sheet = new_book.active

#label Axies and Data
for x in range(1, max_arrow+1):
    new_sheet[get_column_letter(x + 1) + "1"] = x

#get an alphabetized list to iterate through
keys_list = list(meanings_to_lst_arrow.keys())
keys_list = [x for x in keys_list if x != None]
keys_list.sort()
index = 2
for key in keys_list:
    new_sheet["A" + str(index)] = key
    for num in meanings_to_lst_arrow[key]:
        print("Plotting " + str(key) + " to " + str(get_column_letter(num + 1)) + str(index))
        new_sheet[str(get_column_letter(num + 1)) + str(index)].style = Style(fill=myFill)
    index += 1

#save the new workbook with the given output filename
new_book.save(output_name)