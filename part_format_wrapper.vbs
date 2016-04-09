'part_format_wrapper.vbs
'Author: Bill Jameson
'VBScript wrapper for the Python script that formats the data for DMT
'Prerequisite: Inventor-populated data in Part_Level.xls

'TODO:
'if feasible to have Python installed on Inventor boxes, call a Python script to create a CSV file with the data from Inventor (filling in the remaining fields that Epicor requires but will remain constant
'after Python script returns, call DMT at command line w/ Add/Update options