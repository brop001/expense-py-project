import openpyxl
from openpyxl import utils
import re
import numbers

wb = openpyxl.load_workbook(filename='test.xlsx')
ws = wb.active
column_details = ws['C']

config_wb = openpyxl.load_workbook(filename='config.xlsx')
config_ws = config_wb.active


def get_column(column_name, worksheet):
    for cell in worksheet[1]:
        if cell.value == column_name:
            # print(cell.column)
            return worksheet[openpyxl.utils.get_column_letter(cell.column)]
    raise Exception("Can not find the given column!")


def print_column(column):
    for cell in column:
        print(cell.value)


def get_cell_value(row, column, worksheet):
    value = worksheet.cell(row=row, column=column).value
    if isinstance(value, str):
        if value == "None":
            return ""
        else:
            return str(value)
    elif isinstance(value, numbers.Number):
        return int(value)



def get_regex_str(i, cell3):
    regex_mode = get_cell_value(row=i, column=2, worksheet=config_ws)
    if regex_mode == "True":
        regex_string = ".* " + str(cell3.value) + " .*"
    elif regex_mode == "False":
        regex_string = ".*" + str(cell3.value) + ".*"
    else:
        raise Exception("Can not detect regex mode!")
    print("regex string: " + regex_string)
    return regex_string


# Expense regex;Search for the word;Expense description;Expense category;Other category;Expense nature
column_regex = get_column("Expense regex", config_ws)


print_column(column_details)


for idx1, cell1 in enumerate(column_regex, start=1):
    if(idx1 > 1):

        expense_in_Ft = 0
        regex_string = get_regex_str(idx1, cell1)

        expense_description = get_cell_value(row=idx1, column=3, worksheet=config_ws)
        expense_category = get_cell_value(row=idx1, column=4, worksheet=config_ws)
        expense_other_category = get_cell_value(row=idx1, column=5, worksheet=config_ws)
        expense_nature = get_cell_value(row=idx1, column=6, worksheet=config_ws)

        for idx2, cell2 in enumerate(column_details, start=1):
            if re.search(regex_string, str(cell2.value)):
                print("Regex found" + regex_string)
                ws.cell(row=idx2, column=5, value=expense_description)
                ws.cell(row=idx2, column=6, value=expense_category)
                ws.cell(row=idx2, column=7, value=expense_other_category)
                ws.cell(row=idx2, column=8, value=expense_nature)
                expense_in_Ft += get_cell_value(row=idx2, column=2, worksheet=ws)

        ws.cell(row=idx1, column=9, value=regex_string)
        ws.cell(row=idx1, column=10, value=expense_in_Ft)

wb.save('test.xlsx')