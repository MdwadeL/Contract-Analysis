#------------------------------------------------------------------------------------------------
#--Importing the needed libraries and setting the constant variables-----------------------------
#------------------------------------------------------------------------------------------------

# pip install (....)
file = "/Users/macahi.d.wade/Downloads/Projects/Contract Analysis/Contract Master List.xlsx"
sheet = "RawContractAll"

import pandas as pd                 
import matplotlib.pyplot as plt
from time import localtime

cols_to_drop = ['Description', 'POSentDate', 'DueDate', 'TermMonths']
unknown_value_replace = ['BusinessArea']
date_null_columns = ['EndDate']
explore_dates = ['EndDate', 'StartDate']

current_year = localtime().tm_year
current_month = localtime().tm_mon


#------------------------------------------------------------------------------------------------
#--File Locating and Opening---------------------------------------------------------------------
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
        raise

    except ValueError:
        print(f"Error: the sheet [{sheet}] was not found")

        excel_workbook = pd.ExcelFile(file, engine="openpyxl")
        print("Present sheets:", excel_workbook.sheet_names)
        raise


data = open_sheet(sheet)


#------------------------------------------------------------------------------------------------
#--Data Cleaning---------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------

def drop_unneeded_cols(data):
    return data.drop(columns=cols_to_drop, errors='ignore')

def unknown_replaced(data):

    for col in unknown_value_replace:
        data[col] = data[col].fillna('UNKNOWN')
    
    return data

def flag_empty_dates(data):
    for column in date_null_columns:
        data[f'{column}_flag'] = data[column].isnull().astype(int)

    return data

def cleaned_data_csv(data):
    data.to_csv('cleaned_data.csv', index=False)
    return data

data = drop_unneeded_cols(data)
data = unknown_replaced(data)
data = flag_empty_dates(data)
data = cleaned_data_csv(data)


#------------------------------------------------------------------------------------------------
#--Data Featuring--------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------

def date_features(data):
    for column in explore_dates:
        data[f'{column}_month'] = data[column].dt.strftime('%m')
        data[f'{column}_quarter'] = data[column].dt.quarter
        data[f'{column}_year'] = data[column].dt.strftime('%Y')
        data[f'{column}_YrMth'] = data[column].dt.strftime('%Y-%m')
        data[f'{column}_YrQrt'] = data[column].dt.strftime('%Y').astype(str) + '-Q' + data[column].dt.quarter.astype(str)

    return data

def negotiate_savings(data):
    data['NegotiationSavings'] = (data['Quote'] - data['Negotiate']).round(2)

    return data

def savings_percent(data):
    data['SavingsPercent'] = ((
        (data['Quote'] - data['Negotiate']) / data['Quote']
    ) * 100).round(2)

    return data

def feature_data_csv(data):
    data.to_csv('feature_data.csv', index=False)
    return data



data = date_features(data)
data = negotiate_savings(data)
data = savings_percent(data)
data = feature_data_csv(data)

#------------------------------------------------------------------------------------------------
#--Exploratory Data Analysis---------------------------------------------------------------------
#------------------------------------------------------------------------------------------------


"""
Overall contract volume
1. How many contracts exist by status?
2. How many contracts exist by contract type?
3. Which business areas have the most contracts?
4. Which vendors have the highest total contract value?
5. What is the distribution of contract values?
6. How much money is saved through negotiation?
7. Which negotiators produce the greatest negotiation savings?
8. How has contract volume changed by year and quarter?
9. How many contracts are expiring in upcoming months/quarters?
10. Which business areas and negotiators have the largest upcoming expiration workload?
"""








"""
GIT CHANGES

git add icd.py
git commit -m "Change Check"
git push
"""






















"""
GIT CHANGES

git add icd.py
git commit -m "Change Check"
git push
"""
