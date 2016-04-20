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

# Note: the documentation states that xlrd.empty_cell is a singleton,
# but this is false: https://github.com/python-excel/xlrd/issues/162
data = [(field[0].value, field[1].value) for field in zip(title_row, data_row)
        if field[1].value != xlrd.empty_cell.value]
