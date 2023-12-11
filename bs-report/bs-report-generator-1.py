import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# set today date
today = pd.Timestamp.today().strftime('%Y-%m-%d')

# report title
title = 'Employee Task Report - QA Team' 

# excel file name
file_name = "python-playground/bs-report/data/Employee Task Report.xlsx"
new_file_name = f"python-playground/bs-report/data/Employee Task Report {today}.xlsx" 
sheet_name = 'Employee Task'
today = pd.Timestamp.today().strftime('%Y-%m-%d')
new_sheet_name = 'Employee Task ' + today

# read 'Employee Task' sheet of excel file in data folder as dataframe
df_emp_task = pd.read_excel(file_name, sheet_name = sheet_name)
df_emp_task.info()

# set Planned Start, Planned End, Actual Start, Actual End dates as datetime
df_emp_task['Planned Start'] = pd.to_datetime(df_emp_task['Planned Start']).dt.date
df_emp_task['Planned End'] = pd.to_datetime(df_emp_task['Planned End']).dt.date
df_emp_task['Actual Start'] = pd.to_datetime(df_emp_task['Actual Start']).dt.date
df_emp_task['Actual End'] = pd.to_datetime(df_emp_task['Actual End']).dt.date

# select cols to hide from excel sheet
cols_to_hide_list = [
        'Dept and Team', 'Reviewer', 'Requirement Collector', 'Task No'
    ]

# create a new column 'Assign Type'
df_emp_task['Assign Type'] = np.where(df_emp_task['Task Owner'] == df_emp_task['QA'], 'Own Task', 'Assigned Task')

# sort values by QA and Task Owner
df_emp_task = df_emp_task.sort_values(by=['QA', 'Task Owner'])

# Open writer on new workbook
writer = pd.ExcelWriter(new_file_name, engine = 'xlsxwriter')

workbook = writer.book

# Create new sheet
worksheet = workbook.add_worksheet(new_sheet_name) 

# Add the new sheet to the writer
writer.sheets[new_sheet_name] = worksheet

# set header formatting
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'align': 'center',
    'valign': 'vcenter',
    'bg_color': '',
    'font_color': '#FFFFFF'
})

# # merge cells
format = workbook.add_format()
format.set_font_size(25)
format.set_bold()
format.set_font_color('#0000FF')

merged_range = worksheet.merge_range('A1:Z1', title, format)
# set height of row 2 to 15
worksheet.set_row(1, 15) 

# write the column header with defined format
for col_num, value in enumerate(df_emp_task.columns.values):
    worksheet.write(1, col_num, value, header_format)

# total formatting
total_format = workbook.add_format({
    'bold': True,
    'align': 'right',
    'bottom': 6})

# add total rows
number_rows = len(df_emp_task.index)
# Subtract 1 for each merged cell in the title
if merged_range:
    merged_rows = merged_range.split(':')[1].split(':')[0]
    number_rows -= int(merged_rows) - 1

for column in range(14, 17):
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
worksheet.set_column('A:Z', 15)

#############################

# import random
# import colorsys

# # Get the unique QA names
# qa_names = df_emp_task['QA'].unique()

# # Generate random colors with low saturation and brightness
# qa_colors = {}
# for qa_name in qa_names:
#     # Generate a random hue value between 0 and 1
#     hue = random.random()
#     # Set the saturation and brightness values
#     saturation = random.uniform(0.1, 0.3)  # Adjust these values to control the saturation
#     brightness = random.uniform(0.5, 0.6)  # Adjust these values to control the brightness
#     # Convert the HSV color to RGB
#     rgb = colorsys.hsv_to_rgb(hue, saturation, brightness)
#     # Convert the RGB color to hexadecimal format
#     color = '#' + ''.join([f'{int(c * 255):02x}' for c in rgb])
#     qa_colors[qa_name] = color

# # Function to apply conditional formatting
# def apply_color(row):
#     return ['background-color: {}'.format(qa_colors[row['QA']]) for _ in row]

# # Apply the conditional formatting
# styled_df = df_emp_task.style.apply(apply_color, axis=1)

# # Save the styled dataframe to Excel
# styled_df.to_excel(writer, sheet_name=new_sheet_name, index=False)


#############################

# save
writer.close()