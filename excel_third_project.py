import openpyxl
from openpyxl import utils
import re
import os
from datetime import datetime
import pandas


source_data_folder_path = "source_data\\Peti\\"

needed_columns_peti = ["Transaction date", "Transaction amount", "Details"]

file_list = os.listdir(source_data_folder_path)
print(file_list)

def delete_row_with_merged_ranges(sheet, idx):
    begin_col, end_col = [], []
    # Get column information of merged cell of row to delete
    for n in sheet.merged_cells:
        if (n.min_row == idx and n.max_row == idx):
            begin_col.append(n.min_col)
            end_col.append(n.max_col)
    # Cancel cell merge of deleted row
    for n in range(0, len(begin_col), 1):
        sheet.unmerge_cells(
            start_row=idx,
            start_column=begin_col[n],
            end_row=idx,
            end_column=end_col[n]
        )
    sheet.delete_rows(idx)
    for mcr in sheet.merged_cells:
        if idx < mcr.min_row:
            mcr.shift(row_shift=-1)
        elif idx <= mcr.max_row:
            mcr.shrink(bottom=1)


def get_column_by_name(column_name, max_num_of_lines_from_1, worksheet):
    done_flag = False
    for idx in range(1, max_num_of_lines_from_1):
        row = worksheet[idx]
        for cell in row:
            if cell.value == column_name:
                column = worksheet[openpyxl.utils.get_column_letter(cell.column)]
                done_flag = True
                break
        if done_flag:
            break

    if done_flag:
        return column
    else:
        raise Exception("Can not find the given column!")

def delete_row_before_column_name(column_name, column, worksheet):
    done_flag = False
    for idx, cell in enumerate(column, start=1):
        if cell.value == column_name:
            column_name_row_num = idx
            done_flag = True
    if done_flag:
        for idx2 in range(1, column_name_row_num+1):
            print(idx2)
            delete_row_with_merged_ranges(worksheet, idx2)
    else:
        raise Exception("Can not find the given column name!")


def print_column(column):
    for cell in column:
        print(cell.value)


def get_the_newest_file():
    for idx, file in enumerate(file_list, start=1):
        wb_newest = openpyxl.load_workbook(filename=file)
        ws_newest = wb_newest.active

        for idx2 in range(1, 10):
            row_newest = ws_newest[idx2]
            for cell_newest in row_newest:
                pass


        "([0-9]{2})-([0-9]{2})-([0-9]{4})"

        wb_newest.close()


var1 = datetime.now()
var2 = datetime(1993,1,23)

if(var1<var2):
    print(str(var1) + " is newer")
else:
    print(str(var2) + " is newer")

file_path = source_data_folder_path + file_list[0]
print(file_path)

#wb = openpyxl.load_workbook(filename=source_data_folder_path + file_list[0])
#ws = wb.active

#col_name = needed_columns_peti[0]
#print_column(get_column_by_name(col_name, 5, ws))

df = pandas.read_excel(file_path, header=3)

for idx in range(1, len(file_list)):
    print(source_data_folder_path + file_list[idx])
    df2 = pandas.read_excel(source_data_folder_path + file_list[idx], header=3)

    frames = [df, df2]
    df = pandas.concat(frames, ignore_index=True)

df = df[["Transaction date", "Transaction amount", "Details"]]

df.drop_duplicates(subset=["Transaction date", "Transaction amount", "Details"], keep="first", inplace=True)

df.to_excel("output.xlsx")

#wb.save("test2.xlsx")