#!/usr/bin/env python
# coding: utf-8
  
# In[30]:
import pandas as pd
from pandas import option_context
from tabulate import tabulate
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import tkinter as tk
from tkinter import filedialog
import easygui
 
# Hide the main tkinter window
root = tk.Tk()
root.withdraw()

#attempting pop-up box for instructions
easygui.msgbox("Please select the CSV file you would like to convert to an Excel worksheet.", title="simple gui")

# Open file dialog to select the input CSV file
input_file_path = filedialog.askopenfilename(
    title="Select CSV File",
    filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*"))
)
if not input_file_path:
    print("No file selected. Exiting...")
    exit()
 
#attempting pop-up box for instructions
easygui.msgbox("Please select a destination for your new .xlsx output of your .csv file. In order to create a new file, simply type in the desired name, select file destination, and click save.", title="simple gui")

# Open file dialog to select the output Excel file destination
output_file_path = filedialog.asksaveasfilename(
    title="Save Excel File",
    defaultextension=".xlsx",
    filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*"))
)
if not output_file_path:
    print("No save location selected. Exiting...")
    exit()
 
# In[20]:
try:
    df = pd.read_csv(input_file_path)
    print("CSV file loaded successfully.")
except Exception as e:
    print(f"Error loading CSV file: {e}")
    exit()
  
# In[22]:
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
 
# In[10]:
new_df = df.dropna(axis=1, thresh=1)
 
# In[31]:
for row in dataframe_to_rows(new_df, index=False, header=True):
    ws.append(row)
 
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width
 
try:
    wb.save(output_file_path)
    print("Workbook saved successfully.")
except Exception as e:
    print(f"Error saving workbook: {e}")
    exit()