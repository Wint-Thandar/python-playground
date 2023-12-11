import pandas as pd
import numpy as np
from openpyxl import load_workbook
import matplotlib.pyplot as plt
from pathlib import Path

# excel file name
file_name = "python-playground/bs-report/data/Employee Task Report.xlsx"
sheet_name = 'Employee Task'
today = pd.Timestamp.today().strftime('%Y-%m-%d')
new_sheet_name = 'Employee Task ' + today

# read 'Employee Task' sheet of excel file in data folder as dataframe
df_emp_task = pd.read_excel("python-playground/bs-report/data/Employee Task Report.xlsx", sheet_name = sheet_name)
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

# create a new sheet
df_emp_task.to_excel(file_name, sheet_name = new_sheet_name, index=False)

# Apply the formatting condition
df_emp_task.loc[(df_emp_task['Assign Type'] == 'Own Task') & (df_emp_task['Dev Over Due Days'] > 0), 'Dev Over Due Days'] = \
    df_emp_task.loc[(df_emp_task['Assign Type'] == 'Own Task') & (df_emp_task['Dev Over Due Days'] > 0), 'Dev Over Due Days'].apply(
        lambda x: f"<span style='color:red'>{x}</span>"
    )

# Save the modified DataFrame back to the Excel file
with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
    writer.book = load_workbook(file_name)
    df_emp_task.to_excel(writer, sheet_name=new_sheet_name, index=False)
    writer.save()


# resize cols

# hide cols
