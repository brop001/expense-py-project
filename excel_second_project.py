import openpyxl
from openpyxl import utils
import re

wb = openpyxl.load_workbook(filename='test.xlsx')
ws = wb.active
column_details = ws['C']

def print_column(column):
    for cell in column:
        pass
        print(cell.value)

print_column(column_details)

tag_dict = {
  ".*Penny.*": "Bevasarlas",
  ".*Spar.*": "Bevasarlas",
  ".*DM.*": "Háztartási koltsegek"
}

for key, value in tag_dict.items():
    print(key, value)
    for idx, cell in enumerate(column_details, start=1):
        if re.search(key, str(cell.value)):
            ws.cell(row=idx, column=5, value=value)


wb.save('test.xlsx')