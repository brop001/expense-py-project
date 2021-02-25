import openpyxl
from openpyxl import utils
import re
import os
from datetime import datetime
import pandas
from collections import namedtuple
import numpy as np

source_folder_path = "source_data\\"
output_folder_path = "output\\"

combined_file_name = "Combined_output.xlsx"



File_format = namedtuple('File_format', ['file_name', 'columns', 'header_index', 'source_folder', 'columns_to_merge'])

peti = File_format("Peti_output.xlsx", ["Transaction date", "Transaction amount", "Merged_columns"], 3, source_folder_path + "Peti\\",
                   ["Partner name/Secondary account identifier type", "Transaction type", "Details"])
vanda = File_format("Vanda_output.xlsx", ["Tranzakció időpontja", "Összeg", "Merged_columns"], 13, source_folder_path + "Vanda\\",
                    ["Forgalom típusa", "Ellenoldali név", "Közlemény"])


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


def get_column_num_by_name(column_name, max_num_of_lines_from_1, worksheet):
    done_flag = False
    for idx in range(1, max_num_of_lines_from_1):
        row = worksheet[idx]
        for cell in row:
            if cell.value == column_name:
                column = cell.column
                done_flag = True
                break
        if done_flag:
            break

    if done_flag:
        return column
    else:
        raise Exception("Can not find the given column!")


def print_column(column):
    for cell in column:
        print(cell.value)


def create_one_file_form_files(file_format):
    output_file_path = output_folder_path + file_format.file_name

    # Discover files in folder
    file_list = os.listdir(file_format.source_folder)
    files_path_list = []
    for idx, file in enumerate(file_list):
        files_path_list.append(file_format.source_folder + file)

    # Combine all files in the folder in one dataframe
    file_path = files_path_list[0]
    print(file_path)
    df = pandas.read_excel(file_path, header=file_format.header_index)
    for idx in range(1, len(files_path_list)):
        print(files_path_list[idx])
        df2 = pandas.read_excel(files_path_list[idx], header=file_format.header_index)
        frames = [df, df2]
        df = pandas.concat(frames, ignore_index=True)

    # Merge the useful columns into one
    df[file_format.columns[2]] = " "
    for col_name in file_format.columns_to_merge:
        df[file_format.columns[2]] += " " + df[col_name].fillna(" ")

    # Keep only the given columns
    df = df[file_format.columns]

    # Remove duplicated rows
    df.drop_duplicates(subset=file_format.columns, keep="first", inplace=True)

    # Rename all columns to unify them
    df.rename(
        columns={df.columns[0]: "Transaction date", df.columns[1]: "Transaction amount", df.columns[2]: "Details"},
        inplace=True)

    # Save dataframe into excel file
    df.to_excel(output_file_path)

    # Unify time format
    unify_time_format(output_file_path)

    # Reorder columns by date
    order_rows_by_column(output_file_path, "Transaction date")

    return output_file_path


def order_rows_by_column(file_path, column):
    df = pandas.read_excel(file_path, index_col=0)
    pandas.to_datetime(df[column])
    df = df.sort_values(by=column, ascending=False)
    df = df.reset_index(drop=True)

    df.to_excel(file_path)


def unify_time_format(file_path):
    wb = openpyxl.load_workbook(filename=file_path)
    ws = wb.active
    column_details = get_column_by_name("Details", 10, ws)
    column_details_num = get_column_num_by_name("Details", 10, ws)
    column_time_num = get_column_num_by_name("Transaction date", 10, ws)

    for idx in range(2, len(column_details) + 1):
        result = re.search(".*(2[0-9])([0-1][0-9])([0-3][0-9])([0-2][0-9]):([0-6][0-9]).*",
                           str(ws.cell(row=idx, column=column_details_num).value))
        result2 = re.search(".*([0-9]{4}).([0-1][0-9]).([0-3][0-9]) ([0-2][0-9]):([0-6][0-9]):([0-6][0-9]).*",
                           str(ws.cell(row=idx, column=column_details_num).value))
        result3 = re.search(".*([0-1][0-9])-([0-3][0-9])-([0-9]{4}).*",
                            str(ws.cell(row=idx, column=column_time_num).value))
        result4 = re.search(".*([0-9]{4}).([0-1][0-9]).([0-3][0-9]). ([0-2][0-9]):([0-6][0-9]):([0-6][0-9]).*",
                            str(ws.cell(row=idx, column=column_time_num).value))

        if result:
            date = datetime(int(result.group(1)) + 2000, int(result.group(2)), int(result.group(3)),
                            int(result.group(4)), int(result.group(5)))
            # print(date)
        elif result2:
            date = datetime(int(result2.group(1)), int(result2.group(2)), int(result2.group(3)), int(result2.group(4)),
                            int(result2.group(5)), int(result2.group(6)))
            # print(date)
        elif result3:
            date = datetime(int(result3.group(3)), int(result3.group(1)), int(result3.group(2)))
            # print(date)
        elif result4:
            date = datetime(int(result4.group(1)), int(result4.group(2)), int(result4.group(3)), int(result4.group(4)),
                            int(result4.group(5)), int(result4.group(6)))
            # print(date)
        else:
            print(str(ws.cell(row=idx, column=column_details_num).value))
            print(str(ws.cell(row=idx, column=column_time_num).value))
            print(str(ws.cell(row=idx, column=column_time_num).value))
            raise Exception("Can not detect time!")

        ws.cell(row=idx, column=column_time_num, value=str(date))

    wb.save(file_path)


def combine_two_files(file1, file2):
    output_file_path = output_folder_path + combined_file_name

    df1 = pandas.read_excel(file1, index_col=0)
    df2 = pandas.read_excel(file2, index_col=0)

    frames = [df1, df2]
    df = pandas.concat(frames, ignore_index=True)

    df.to_excel(output_file_path)

    order_rows_by_column(output_file_path, "Transaction date")

    return output_file_path


def separate_file_by_date(file_path):
    month_list = get_date_list(file_path, False)
    year_list = get_date_list(file_path, True)

    df = pandas.read_excel(file_path, index_col=0)
    for date in month_list:
        df_filter = df["Transaction date"].str.contains(".*" + date + ".*")
        df2 = df[df_filter]
        df2 = df2.reset_index(drop=True)
        df2.to_excel(output_folder_path + date + "_output.xlsx")

    for date in year_list:
        df_filter = df["Transaction date"].str.contains(".*" + date + ".*")
        df2 = df[df_filter]
        df2 = df2.reset_index(drop=True)
        df2.to_excel(output_folder_path + date + "_output.xlsx")


def get_date_list(file_path, only_year):
    wb = openpyxl.load_workbook(filename=file_path)
    ws = wb.active
    column_time = get_column_by_name("Transaction date", 10, ws)
    column_time_num = get_column_num_by_name("Transaction date", 10, ws)

    date_list = []

    for idx in range(2, len(column_time) + 1):

        result = re.search("([0-9]{4})-([0-1][0-9])", str(ws.cell(row=idx, column=column_time_num).value))
        if result:
            if only_year:
                date = str(result.group(1))
            else:
                date = str(result.group(0))
            if not date_list.count(date):
                date_list.append(date)
    return date_list


def add_more_columns(folder_path):
    # Discover files in folder
    file_list = os.listdir(folder_path)
    files_path_list = []
    for idx, file in enumerate(file_list):
        files_path_list.append(folder_path + file)


output_file_path_peti = create_one_file_form_files(peti)
output_file_path_vanda = create_one_file_form_files(vanda)

output_combined_file_path = combine_two_files(output_file_path_peti, output_file_path_vanda)

separate_file_by_date(output_combined_file_path)
add_more_columns(output_folder_path)
