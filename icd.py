"""
Importing the needed libraries and variables
"""

# pip install (....)
import pandas as pd                 
import matplotlib.pyplot as plt
import numpy as np
from time import localtime

year = localtime().tm_year
month = localtime().tm_mon

folder = "/Users/macahi.d.wade/Downloads/Projects/Contract Analysis"
path = f"Contract Primary Report {year}-{month:02}.xlsx"
rca = "RawContractAll"
file = folder + "/" + path


"""
File Locating and Opening
"""

def open_sheet(sheet):
    try:
        return pd.read_excel(file, sheet_name=rca, engine="openpyxl")
    except FileNotFoundError:
        print(f"Error: the path [{file}] was not found")
    except ValueError:
        print(f"Your sheet [{rca}] was not found")

        excel_workbook = pd.ExcelFile(file, engine="openpyxl")
        print("Present sheets: ", excel_workbook.sheet_names)

raw_contract = open_sheet(rca)

"""
Data Cleaning
"""

print('hi there')



"""
GIT CHANGES

git add .
git commit -m "Change Check"
git push
"""