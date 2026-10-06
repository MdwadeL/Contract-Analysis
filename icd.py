#------------------------------------------------------------------------------------------------
#--Importing the needed libraries and setting the constant variables
#------------------------------------------------------------------------------------------------

# pip install (....)
file = "/Users/macahi.d.wade/Downloads/Projects/Contract Analysis/Contract Master List.xlsx"
sheet = "RawContractAll"

import pandas as pd                 
import matplotlib.pyplot as plt
import numpy as np
from time import localtime

cols_to_drop = ['Description', 'POSentDate', 'DueDate', 'TermMonths']
unknown_value_replace = ['BusinessArea']

current_year = localtime().tm_year
current_month = localtime().tm_mon


#------------------------------------------------------------------------------------------------
#--File Locating and Opening
#------------------------------------------------------------------------------------------------

def open_sheet(sheet):
    try:
        return pd.read_excel(
            file,
            sheet_name=sheet,
            engine="openpyxl"
        )

    except FileNotFoundError:
        print(f"Error: the file [{file}] was not found")

    except ValueError:
        print(f"Error: the sheet [{sheet}] was not found")

        excel_workbook = pd.ExcelFile(file, engine="openpyxl")
        print("Present sheets:", excel_workbook.sheet_names)


data = open_sheet(sheet)

#------------------------------------------------------------------------------------------------
#--Data Cleaning
#------------------------------------------------------------------------------------------------

def drop_unneeded_cols(data):

    for column in cols_to_drop:
        if column not in data.columns:
            #print(f'{column} was not found in columns to drop')
            continue
        
        else: data = data.drop(columns=column)

    return data

def unknown_replaced(data):

    for col in unknown_value_replace:
        data[col] = data[col].replace(pd.NA, 'Unknown')
    
    return data



data = drop_unneeded_cols(data)
data = unknown_replaced(data)






















"""
GIT CHANGES

git add icd.py
git commit -m "Change Check"
git push
"""