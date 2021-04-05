import xlrd
import xlsxwriter
import csv
import numpy as np

# Import CSV File and make it 2D array
print("Please enter your CSV path")
path = input()
print("Enter name of CSV file: ")
name_of_file = input()
raw_string = r"{}".format(path)
current_path = f"{raw_string}\\{name_of_file}.csv"
csv_file = open(current_path)
reader = csv.reader(csv_file)


csv_data = []
for row in reader:
    csv_data.append(row)

# print(csv_data)

# Import XLSX File and make it 2D array
print("Please enter your XLSX path")
path = input()
print("Enter name of XLSX file: ")
name_of_file = input()
raw_string = r"{}".format(path)
current_path = f"{raw_string}\\{name_of_file}.xlsx"

workbook = xlrd.open_workbook(current_path)
worksheet = workbook.sheet_by_index(0)

xlsx_data = []
# Creating 2D Array and fill it with all xlsx information
for row in range(worksheet.nrows):
    current_row = []
    for col in range(worksheet.ncols):
        current_row.append(worksheet.cell_value(row, col))
    xlsx_data.append(current_row)

print(xlsx_data)

# Splitting data from CSV and XLSX file into one 2D Array and Validate the data
for i in range(len(xlsx_data)):
    if type(xlsx_data[i][0]) == str and type(csv_data[i][1]) == str:
        xlsx_data[i][0] = f'{xlsx_data[i][0]}   {csv_data[i][1]}'
    elif type(xlsx_data[i][0]) == float and type(csv_data[i][1]) == float:
        xlsx_data[i][0] = f'{xlsx_data[i][0]}   {csv_data[i][1]}'

    if type(xlsx_data[i][1]) == str and type(csv_data[i][2]) == str:
        xlsx_data[i][1] = f'{xlsx_data[i][1]}   {csv_data[i][2]}'
    elif type(xlsx_data[i][0]) == float and type(csv_data[i][2]) == float:
        xlsx_data[i][1] = f'{xlsx_data[i][1]}   {csv_data[i][2]}'

    if type(xlsx_data[i][2]) == str and type(csv_data[i][3]) == str:
        xlsx_data[i][2] = f'{xlsx_data[i][2]}   {csv_data[i][3]}'
    elif type(xlsx_data[i][2]) == float and type(csv_data[i][3]) == float:
        xlsx_data[i][2] = f'{xlsx_data[i][2]}    {csv_data[i][3]}'

    if type(xlsx_data[i][3]) == str and type(csv_data[i][4]) == str:
        xlsx_data[i][3] = f'{xlsx_data[i][3]}    {csv_data[i][4]}'
    elif type(xlsx_data[i][3]) == float and type(csv_data[i][4]) == float:
        xlsx_data[i][3] = f'{xlsx_data[i][3]}    {csv_data[i][4]}'

    if type(xlsx_data[i][4]) == str and type(csv_data[i][0]) == str:
        xlsx_data[i][4] = f'{xlsx_data[i][4]}    {csv_data[i][0]}'
    elif type(xlsx_data[i][4]) == float and type(csv_data[i][0]) == float:
        xlsx_data[i][4] = f'{xlsx_data[i][4]}    {csv_data[i][0]}'

#print(xlsx_data)

raw_string = r"{}".format(path)
current_path = f"{raw_string}\\Modified_{name_of_file}.xlsx"
new_workbook = xlsxwriter.Workbook(current_path)
new_sheet = new_workbook.add_worksheet()

# Filling New XLSX File with all the information
for row in range(len(xlsx_data)):
    for col in range(len(xlsx_data[0])):
        new_sheet.write(row, col, xlsx_data[row][col])

new_workbook.close()
