import pandas as pd
import numpy as np
import openpyxl
from xlsxwriter.utility import xl_rowcol_to_cell
import matplotlib.pyplot as plt
from pathlib import Path

# excel file name
file_name = "python-playground/bs-report/data/Employee Task Report.xlsx"
sheet_name = 'Employee Task'
today = pd.Timestamp.today().strftime('%Y-%m-%d')
new_sheet_name = 'Employee Task ' + today

# read 'Employee Task' sheet of excel file in data folder as dataframe
df_emp_task = pd.read_excel(file_name, sheet_name = sheet_name)
df_emp_task.info()

# set Planned Start, Planned End, Actual Start, Actual End dates as datetime
df_emp_task['Planned Start'] = pd.to_datetime(df_emp_task['Planned Start'])
df_emp_task['Planned End'] = pd.to_datetime(df_emp_task['Planned End'])
df_emp_task['Actual Start'] = pd.to_datetime(df_emp_task['Actual Start'])
df_emp_task['Actual End'] = pd.to_datetime(df_emp_task['Actual End'])

# select cols to hide from excel sheet
cols_to_hide_list = [
        'Dept and Team', 'Reviewer', 'Requirement Collector', 'Task No'
    ]

# create a new column 'Assign Type'
df_emp_task['Assign Type'] = np.where(df_emp_task['Task Owner'] == df_emp_task['QA'], 'Own Task', 'Assigned Task')

# sort values by QA and Task Owner
df_emp_task = df_emp_task.sort_values(by=['QA', 'Task Owner'])

# get access to worksheet object
writer = pd.ExcelWriter(file_name, engine='xlsxwriter', mode='a')
df_emp_task.to_excel(writer, sheet_name = new_sheet_name, index = False)

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

# merge cells
format = workbook.add_format()
format.set_font_size(20)
format.set_font_color('#333333')

subheader = 'QA Team'
worksheet.merge_range('A1:Z1', title, format)
worksheet.merge_range('A2:Z2', subheader)
worksheet.set_row(2, 15)

# write the column header with defined format
for col_num, value in enumerate(df_emp_task.columns.values):
    worksheet.write(2, col_num, value, header_format)

# total formatting
total_format = workbook.add_format({
    'bold': True,
    'align': 'right',
    'bottom': 6})

# add total rows
number_rows = len(df_emp_task.index)
for column in range(14, 19):
    # determine where to write the total
    cell_location = xl_rowcol_to_cell(number_rows+1, column)
    
    # get the range to use for the sum formula
    start_range = xl_rowcol_to_cell(3, column)
    end_range = xl_rowcol_to_cell(number_rows, column)
    
    # set formula
    formula = '=SUM({:s}:{:s})'.format(start_range, end_range)
    
    # write the formula
    worksheet.write_formula(cell_location, formula, total_format)

# adjust the column width
worksheet.set_column('A:J', 20)

# save
writer._save()