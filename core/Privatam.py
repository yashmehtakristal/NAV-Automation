import msoffcrypto
import xlrd
import openpyxl
import os
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pickle
import shutil
import streamlit as st


def rename_files_with_prefix(directory_path, prefix, new_prefix, extensions):
    """
    Rename files in the specified directory that start with a given prefix and have specific extensions.

    Args:
        directory_path (str): The path to the directory where the files are located.
        prefix (str): The prefix that files should start with to be considered for renaming.
        new_prefix (str): The new prefix to be used in the renamed files.
        extensions (list): A list of file extensions to consider for renaming.

    Returns:
        excel_file_path_ubs (list): List of the .xlsx files (main data source) of Privatam
    """
    excel_file_path_privatam = []
    files_in_directory = os.listdir(directory_path)

    for file_name in files_in_directory:
        if file_name.startswith(prefix) and any(file_name.lower().endswith(ext) for ext in extensions):
            date_part = file_name[len(prefix):len(prefix) + 10]
            formatted_date = f"{date_part[8:10]}-{date_part[5:7]}-{date_part[0:4]}"
            new_file_name = f"{new_prefix}_{formatted_date}{os.path.splitext(file_name)[1]}"
            old_file_path = os.path.join(directory_path, file_name).replace("\\", "/")
            new_file_path = os.path.join(directory_path, new_file_name).replace("\\", "/")
            if new_file_path.endswith('.xlsx'):
                excel_file_path_privatam.append(new_file_path)

            shutil.copy2(old_file_path, new_file_path)
            # print(f"Successfully renamed {file_name} to {new_file_name}")
            
    return excel_file_path_privatam


def keep_only_products_worksheet(input_file_path):
    """
    Keep only the 'Products' worksheet in the Excel workbook and remove all other worksheets.

    Args:
        input_file_path (str): The path to the input Excel file.

    Returns:
        None
    """
    workbook = openpyxl.load_workbook(input_file_path)
    sheets_to_remove = [sheet.title for sheet in workbook.worksheets if sheet.title != 'Products']

    for sheet_name in sheets_to_remove:
        workbook.remove(workbook[sheet_name])

    workbook.save(input_file_path)

    print(f"'Products' worksheet saved to {input_file_path}")


def process_excel_privatam(directory_path_privatam, input_file_path):
    """
    Processes the input CSV file and updates it according to desired output specification.

    Parameters:
        input_file_path (str): The path to the input CSV file that needs to be processed

    Returns:
        None - file is updated in-place
    """
    
    data = pd.read_excel(input_file_path)
    # st.dataframe(data)
    header_row_idx = None

    for idx, row in data.iterrows():
        if 'ISIN' in row.values:
            header_row_idx = idx
            break

    data = data.loc[header_row_idx:].reset_index(drop=True)
    data.columns = data.iloc[0]
    data = data.drop(0)

    last_part = input_file_path.split('/')[-1]
    new_last_part = last_part.replace('.csv','.xlsx')
    
    date_str = last_part.split('_')[1].replace('.xlsx', '')
    source_str = last_part.split('_')[0]
    
    new_data = pd.DataFrame({
        "ISIN": data["ISIN"],
        "Price": data["Market Price"],
        "Date": date_str,
        "Source": source_str, # Take the second part of "Privatam/Privatam_" or simply do "Source": new_prefix
        "Issuer": data["Issuer"]
    })
    
    nan_price_isins = new_data[pd.isna(new_data['Price'])]['ISIN'].tolist()
    
    if len(nan_price_isins) > 0:
        print(f"These ISINs don't have any corresponding prices, we have skipped over these: {nan_price_isins}")
        new_data = new_data.dropna(subset=['Price'])  
    
    new_data.dropna()
    output_file_path = directory_path_privatam + f"/final_{new_last_part}"
    new_data.to_excel(output_file_path, index=False)
        
    # print(f"Successfully created output file in {input_file_path}")

    return output_file_path
