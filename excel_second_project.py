import openpyxl
from openpyxl import utils
import re
import numbers
import pandas
import numpy as np
import matplotlib.pyplot as plt
import os

output_folder_path = "output\\"
file_name = "Combined_output.xlsx"

file_path = output_folder_path + file_name
config_path = "config.xlsx"

wb = openpyxl.load_workbook(filename=file_path)
ws = wb.active
column_details = ws['D']

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


def get_regex_str(idx, dataframe):
    regex_str = dataframe.at[idx, "Expense regex"]
    regex_mode = dataframe.at[idx, "Search for the word"]
    if regex_mode == True:
        regex_string = ".* " + regex_str + " .*"
    elif regex_mode == False:
        regex_string = ".*" + regex_str + ".*"
    else:
        raise Exception("Can not detect regex mode!")
    print("regex string: " + regex_string)
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
    print(value_list)
    return value_list


def get_and_write_category_amount(exp_category, exp_category_amount, file_path, dataframe_results, dataframe):
    exp_cat_list = get_value_list(file_path, exp_category)

    for idx1, value in enumerate(exp_cat_list):
        expense_in_Ft = 0

        for idx2, row in dataframe.iterrows():
            if dataframe.at[idx2, exp_category] == value:
                expense_in_Ft += dataframe.at[idx2, "Transaction amount"]

        dataframe_results.at[idx1, exp_category] = value
        dataframe_results.at[idx1, exp_category_amount] = expense_in_Ft
        #dataframe_results = dataframe_results.append({exp_category: value, exp_category_amount: expense_in_Ft}, ignore_index=True)

    values = dataframe_results[exp_category_amount].abs().dropna()
    labels = dataframe_results[exp_category].dropna()
    print(values)
    print(labels)

    plt.pie(values, labels=labels)
    #plt.show()

    return dataframe_results


def write_chart(file_path, chart_list):

    writer = pandas.ExcelWriter(file_path, engine='xlsxwriter')
    df_results.to_excel(writer, sheet_name='Sheet1')

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    for chart in range(len(chart_list)):
        cat = df_results[chart_list[chart][0]]
        val = df_results[chart_list[chart][1]]

        cat_first_col = df_results.columns.get_loc(chart_list[chart][0])+1
        cat_first_row = cat.index.get_loc(cat.first_valid_index())+1
        cat_last_col = df_results.columns.get_loc(chart_list[chart][0])+1
        cat_last_row = cat.index.get_loc(cat.last_valid_index())+1

        val_first_col = df_results.columns.get_loc(chart_list[chart][1])+1
        val_first_row = val.index.get_loc(val.first_valid_index())+1
        val_last_col = df_results.columns.get_loc(chart_list[chart][1])+1
        val_last_row = val.index.get_loc(val.last_valid_index())+1

        pie_chart = (workbook.add_chart({'type': chart_list[chart][3]}))

        pie_chart.add_series({
            'name': chart_list[chart][0],
            'categories': ['Sheet1', cat_first_row, cat_first_col, cat_last_row, cat_last_col],
            'values': ['Sheet1', val_first_row, val_first_col, val_last_row, val_last_col],
            'data_labels': {'value': True, 'percentage': True},
        })

        # Add a title.
        pie_chart.set_title({'name': chart_list[chart][0]})

        # Insert the chart into the worksheet.
        worksheet.insert_chart(chart_list[chart][2], pie_chart, {'x_scale': 3, 'y_scale': 1})

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()


# Expense regex;Search for the word;Expense description;Expense category;Other category;Expense nature

df = pandas.read_excel(file_path, index_col=0)
df_config = pandas.read_excel(config_path, index_col=0)
print(df)
print(df_config)


df["Expense description"] = ""
df["Expense category"] = ""
df["Expense other category"] = ""
df["Expense nature"] = ""

print(df)

df_results = pandas.DataFrame(columns=["Expense regex", "Expense amount",
                                       "Expense category", "Expense category amount",
                                       "Expense other category", "Expense other category amount",
                                       "Expense nature", "Expense nature amount"])

print(df_results)


for idx1, config in df_config.iterrows():
    expense_in_Ft = 0
    regex_string = get_regex_str(idx1, df_config)

    expense_description = df_config.at[idx1, "Expense description"]
    expense_category = df_config.at[idx1, "Expense category"]
    expense_other_category = df_config.at[idx1, "Expense other category"]
    expense_nature = df_config.at[idx1, "Expense nature"]

    for idx2, row2 in df.iterrows():

        if re.search(regex_string, df.at[idx2, "Details"]):
            print("Regex found" + regex_string)
            df.at[idx2, "Expense description"] = expense_description
            df.at[idx2, "Expense category"] = expense_category
            df.at[idx2, "Expense other category"] = expense_other_category
            df.at[idx2, "Expense nature"] = expense_nature
            expense_in_Ft += df.at[idx2, "Transaction amount"]

    df_results = df_results.append({"Expense regex": regex_string, "Expense amount": expense_in_Ft}, ignore_index=True)


df_results = get_and_write_category_amount("Expense category", "Expense category amount", config_path, df_results, df)
df_results = get_and_write_category_amount("Expense other category", "Expense other category amount", config_path, df_results, df)
df_results = get_and_write_category_amount("Expense nature", "Expense nature amount", config_path, df_results, df)

df.to_excel(file_path)

df_results.to_excel(output_folder_path + "Combined_output_results.xlsx")

pie_chart_list = [["Expense category", "Expense category amount", "K2", 'column'],
                  ["Expense other category", "Expense other category amount", "K17", 'pie'],
                  ["Expense nature", "Expense nature amount", "K32", 'pie']]

cell_list = ["K2", "K17", "K32"]

write_chart(os.path.join(output_folder_path, "Combined_output_results.xlsx"), pie_chart_list)