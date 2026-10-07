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

def count_status(data):
    status_counts = data.groupby('Status').agg(count_statuses=('Status', 'count'))
    print(f'\nCount of Statuses')
    print(status_counts)

    return status_counts

def count_contracttype(data):
    contracttype_counts = data.groupby('ContractType').agg(count_contracttypes=('ContractType', 'count'))
    print(f'\nCounts of ContractTypes')
    print(contracttype_counts)

    return contracttype_counts

def count_businessarea(data):
    businessarea_counts = data.groupby('BusinessArea').agg(count_businessarea=('BusinessArea', 'count'))
    print(f'\nCounts of BusinessArea')
    print(businessarea_counts)

    return businessarea_counts

def stats_amounts_vendor(data):
    vendor_amount_stat = data.groupby('Vendor').agg(
        min=('ContractReportCurrentAmount', 'min'),
        average=('ContractReportCurrentAmount', 'mean'),
        max=('ContractReportCurrentAmount', 'max'),
        sum=('ContractReportCurrentAmount', 'sum')
    ).sort_values(by='sum', ascending=False)

    vendor_amount_stat['average'] = vendor_amount_stat['average'].round(2)

    print(f'\nStatistics of Vendor')
    print(vendor_amount_stat)

    return vendor_amount_stat

def stats_amounts_businessarea(data):
    businessarea_amount_stat = data.groupby('BusinessArea').agg(
        min=('ContractReportCurrentAmount', 'min'),
        average=('ContractReportCurrentAmount', 'mean'),
        max=('ContractReportCurrentAmount', 'max'),
        sum=('ContractReportCurrentAmount', 'sum')
    ).sort_values(by='sum', ascending=False)

    businessarea_amount_stat['average'] = businessarea_amount_stat['average'].round(2)

    print(f'\nStatistics Contract Report Current Amount by Business Area')
    print(businessarea_amount_stat)

    return businessarea_amount_stat

def stats_savings_negotiator(data):
    negoitator_savings_stat = data.groupby('Negotiator').agg(
        min=('NegotiationSavings', 'min'),
        average=('NegotiationSavings', 'mean'),
        max=('NegotiationSavings', 'max'),
        sum=('NegotiationSavings', 'sum')
    ).sort_values(by='sum', ascending=False)

    negoitator_savings_stat['average'] = negoitator_savings_stat['average'].round(2)

    print(f'\nStatistics of Savings by Negoitator')
    print(negoitator_savings_stat)

    return negoitator_savings_stat    

status_counts = count_status(data)
contracttype_counts = count_contracttype(data)
businessarea_counts = count_businessarea(data)
vendor_amount_stats = stats_amounts_vendor(data)
businessarea_amount_stats = stats_amounts_businessarea(data)
negotiator_savings_stats = stats_savings_negotiator(data)
