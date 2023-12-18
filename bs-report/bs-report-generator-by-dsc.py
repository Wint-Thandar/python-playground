
# Here is your updated code:

import pandas as pd
import numpy as np
from datetime import date

# excel file name
file_name = "python-playground/bs-report/data/Employee Task Report.xlsx"
sheet_name = 'Employee Task'
today = date.today().strftime('%Y-%m-%d')
new_sheet_name = f'Employee Task {today}'

# read 'Employee Task' sheet of excel file in data folder as dataframe
df_emp_task = pd.read_excel(file_name, sheet_name=sheet_name)
df_emp_task.info()

# set Planned Start, Planned End, Actual Start, Actual End dates as datetime
date_columns = ['Planned Start', 'Planned End', 'Actual Start', 'Actual End']
for column in date_columns:
    df_emp_task[column] = pd.to_datetime(df_emp_task[column])

# select cols to hide from excel sheet
cols_to_hide_list = [
    'Dept and Team', 'Reviewer', 'Requirement Collector', 'Task No'
]

# create a new column 'Assign Type'
df_emp_task['Assign Type'] = np.where(df_emp_task['Task Owner'] == df_emp_task['QA'], 'Own Task', 'Assigned Task')

# sort values by QA and Task Owner
df_emp_task = df_emp_task.sort_values(by=['QA', 'Task Owner'])

# get access to worksheet object
with pd.ExcelWriter(file_name, mode='a') as writer:  # open file in append mode
    df_emp_task.to_excel(writer, sheet_name=new_sheet_name, index=False)

workbook = writer.book
worksheet = writer.sheets[new_sheet_name]

# set header formatting
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '#951F06',
    'font_color': '#FFFFFF'
})

# add title
title = 'Employee Task Report'
subheader = 'QA Team'
worksheet.merge_range('A1:Z1', title, header_format)
worksheet.merge_range('A2:Z2', subheader)

# write the column headers with defined format
for col_num, value in enumerate(df_emp_task.columns.values):
    worksheet.write(3, col_num, value, header_format)

# adjust the column width
worksheet.set_column('A:Z', 20)

# hide columns
for col in cols_to_hide_list:
    worksheet.set_column(col, None, None, {'hidden': True})

# This code will write to a new sheet of the Excel file instead of overwriting the existing sheet. The new sheet is named according to today's date. Note that this code assumes you have pandas and openpyxl installed in your Python environment. If not, use pip install command to install them:

# pip install pandas openpyxl
