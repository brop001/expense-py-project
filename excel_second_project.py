import sys

import openpyxl
from openpyxl import utils
import re
import numbers
import pandas
import numpy as np
from collections import namedtuple
import os

output_folder = "output"
combined_output_xlsx = "Combined_output.xlsx"
peti_output_xlsx = "Peti_output"
vanda_output_xlsx = "Vanda_output"
combined_output_path = os.path.join(output_folder, combined_output_xlsx)
peti_output_path = os.path.join(output_folder, peti_output_xlsx)
vanda_output_path = os.path.join(output_folder, vanda_output_xlsx)

config_path = "config.xlsx"

Chart_config = namedtuple('File_format', ['category_col_name', 'category_value_col_name', 'start_cell', 'chart_type',
                                          'scale_X', 'scale_Y'])
str_exp_regex = "Expense regex"
str_exp_amount = "Expense amount"
str_exp_cat = "Expense category"
str_exp_cat_amount = "Expense category amount"
str_exp_o_cat = "Expense other category"
str_exp_o_cat_amount = "Expense other category amount"
str_exp_nature = "Expense nature"
str_exp_nature_amount = "Expense nature amount"
str_information = "Information"
str_information_value = "Information value"
str_exp_description = "Expense description"
str_status = "Status"

expense_category_ntpl = Chart_config(str_exp_cat, str_exp_cat_amount, "M2", 'column', 3, 1)
expense_other_category_ntpl = Chart_config(str_exp_o_cat, str_exp_o_cat_amount, "M17", 'pie', 1, 1)
expense_nature_ntpl = Chart_config(str_exp_nature, str_exp_nature_amount, "M32", 'pie', 1, 1)
information_ntpl = Chart_config(str_information, str_information_value, "U17", 'pie', 1, 1)


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


def get_regex_str(idx, dataframe):
    regex_str = dataframe.at[idx, str_exp_regex]
    regex_mode = dataframe.at[idx, "Search for the word"]
    if regex_mode == True:
        regex_string = ".* " + regex_str + " .*"
    elif regex_mode == False:
        regex_string = ".*" + regex_str + ".*"
    else:
        raise Exception("Can not detect regex mode!")
    # print("regex string: " + regex_string)
    return regex_string


def get_value_list(file_path, column):
    df = pandas.read_excel(file_path, index_col=0)

    value_list = []

    for idx, row in df.iterrows():
        value = df.at[idx, column]
        if not value_list.count(value):
            value_list.append(value)
    if value_list.count(np.nan):
        value_list.remove(np.nan)
    return value_list


def get_and_write_category_amount(exp_category, exp_category_amount, file_path, dataframe_results, dataframe):
    exp_cat_list = get_value_list(file_path, exp_category)

    for idx1, value in enumerate(exp_cat_list):
        expense_in_Ft = 0

        for idx2, row in dataframe.iterrows():
            if dataframe.at[idx2, exp_category] == value:
                expense_in_Ft += dataframe.at[idx2, "Transaction amount"]

        dataframe_results.at[idx1, exp_category] = value
        dataframe_results.at[idx1, exp_category_amount] = expense_in_Ft*(-1)

    return dataframe_results


def write_chart(file_path, chart_config):

    df_res = pandas.read_excel(file_path, index_col=0)

    writer = pandas.ExcelWriter(file_path, engine='xlsxwriter')
    df_res.to_excel(writer, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    for chart_idx in range(len(chart_config)):
        cat = df_res[chart_config[chart_idx].category_col_name]
        val = df_res[chart_config[chart_idx].category_value_col_name]

        cat_first_col = df_res.columns.get_loc(chart_config[chart_idx].category_col_name)+1
        cat_first_row = cat.index.get_loc(cat.first_valid_index())+1
        cat_last_col = df_res.columns.get_loc(chart_config[chart_idx].category_col_name)+1
        cat_last_row = cat.index.get_loc(cat.last_valid_index())+1

        val_first_col = df_res.columns.get_loc(chart_config[chart_idx].category_value_col_name)+1
        val_first_row = val.index.get_loc(val.first_valid_index())+1
        val_last_col = df_res.columns.get_loc(chart_config[chart_idx].category_value_col_name)+1
        val_last_row = val.index.get_loc(val.last_valid_index())+1

        pie_chart = (workbook.add_chart({'type': chart_config[chart_idx].chart_type}))

        pie_chart.add_series({
            'name': chart_config[chart_idx].category_col_name,
            'categories': ['Sheet1', cat_first_row, cat_first_col, cat_last_row, cat_last_col],
            'values': ['Sheet1', val_first_row, val_first_col, val_last_row, val_last_col],
            'data_labels': {'value': True, 'percentage': True},
        })

        # Add a title.
        pie_chart.set_title({'name': chart_config[chart_idx].category_col_name})

        # Insert the chart into the worksheet.
        worksheet.insert_chart(chart_config[chart_idx].start_cell, pie_chart,
                               {'x_scale': chart_config[chart_idx].scale_X, 'y_scale': chart_config[chart_idx].scale_Y})

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()


def generate_expense_category_labels(file_path):
    print()
    print("Generating labels and charts on this file: " + os.path.basename(file_path))
    print(file_path)
    df = pandas.read_excel(file_path, index_col=0)
    df_config = pandas.read_excel(config_path, index_col=0)

    df[str_exp_description] = ""
    df[str_exp_cat] = ""
    df[str_exp_o_cat] = ""
    df[str_exp_nature] = ""
    df[str_status] = ""

    df_results = pandas.DataFrame(columns=[str_exp_regex, str_exp_amount,
                                           str_exp_cat, str_exp_cat_amount,
                                           str_exp_o_cat, str_exp_o_cat_amount,
                                           str_exp_nature, str_exp_nature_amount,
                                           str_information, str_information_value])
    num_of_processed_rows = 0
    for idx1, config in df_config.iterrows():
        print("Progress: {}%".format(int((idx1 / len(df_config.index)) * 100)), end='\r')
        expense_in_ft = 0

        regex_string = get_regex_str(idx1, df_config)

        expense_description = df_config.at[idx1, str_exp_description]
        expense_category = df_config.at[idx1, str_exp_cat]
        expense_other_category = df_config.at[idx1, str_exp_o_cat]
        expense_nature = df_config.at[idx1, str_exp_nature]

        for idx2, row2 in df.iterrows():
            if re.search(regex_string, df.at[idx2, "Details"]):
                df.at[idx2, str_exp_description] = expense_description
                df.at[idx2, str_exp_cat] = expense_category
                df.at[idx2, str_exp_o_cat] = expense_other_category
                df.at[idx2, str_exp_nature] = expense_nature
                if not df.at[idx2, str_status] == "progressed":
                    df.at[idx2, str_status] = "progressed"
                    num_of_processed_rows += 1
                else:
                    raise Exception("Multiple found by regex on one row! Need investigation in row " + str(idx2))
                expense_in_ft += df.at[idx2, "Transaction amount"]

        df_results = df_results.append({str_exp_regex: regex_string, str_exp_amount: expense_in_ft*(-1)},
                                       ignore_index=True)

    df_results.at[0, str_information] = "Nem feldolgozott sorok száma"
    df_results.at[1, str_information] = "Feldolgozott sorok száma"
    df_results.at[0, str_information_value] = len(df.index) - num_of_processed_rows
    df_results.at[1, str_information_value] = num_of_processed_rows

    df_results = get_and_write_category_amount(str_exp_cat, str_exp_cat_amount, config_path, df_results, df)
    df_results = get_and_write_category_amount(str_exp_o_cat, str_exp_o_cat_amount, config_path, df_results, df)
    df_results = get_and_write_category_amount(str_exp_nature, str_exp_nature_amount, config_path, df_results, df)

    df.to_excel(file_path)
    result_file_name = os.path.basename(os.path.splitext(file_path)[0]) + "_result.xlsx"
    result_file_path = os.path.join(os.path.dirname(file_path), result_file_name)
    df_results.to_excel(result_file_path)

    print("Expenses labelled succesfully! label rate: {}%".format(int((num_of_processed_rows/len(df.index))*100)))

    chart_config_list = [expense_category_ntpl, expense_other_category_ntpl, expense_nature_ntpl, information_ntpl]
    write_chart(result_file_path, chart_config_list)
    print("Charts inserted")


listOfFiles = list()

for (dirpath, dirnames, filenames) in os.walk("output"):
    listOfFiles += [os.path.join(dirpath, file) for file in filenames]

for file in listOfFiles:
    generate_expense_category_labels(file)

