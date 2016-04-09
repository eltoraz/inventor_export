'''part_format.py - consolidate Inventor part data in a more Epicor-friendly format
Author: Bill Jameson
Third-party dependencies: xlrd (https://pypi.python.org/pypi/xlrd)
'''
import xlrd

# TODO:
# - open XLS, grab relevant data, construct CSV
# - check that the required fields have been filled
# - validate input
# - log activity

filename = 'Part_Level.xls'
worksheet = xlrd.open_workbook(filename).sheet_by_index(0)
title_row = worksheet.row(0)
data_row = worksheet.row(1)

# every empty cell should be the same object as xlrd.empty_cell,
# but comparisons still return False
data = [field for field in zip(title_row, data_row)
        if field[1].value != xlrd.empty_cell.value]
