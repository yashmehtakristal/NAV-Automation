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

def getting_file_paths_bloomberg(directory_path_bloomberg, xls_file_path_bloomberg):

    prefix = directory_path_bloomberg.split('/')[-1] # Obtaining the word/prefix "Bloomberg"
    files_in_directory = os.listdir(directory_path_bloomberg)

    for file_name in files_in_directory:
        if file_name.startswith(prefix) and file_name.endswith('.xlsx'):
            full_file_path = os.path.join(directory_path_bloomberg, file_name).replace("\\","/")
            xls_file_path_bloomberg.append(full_file_path)

    return xls_file_path_bloomberg


def process_excel_bloomberg(directory_path_bloomberg, input_file):
    """
    Processes the input CSV file and updates it according to desired output specification.

    Parameters:
        input_file (str): The path to the input CSV file that needs to be processed

    Returns:
        None - file is updated in-place
    """

    file_copy = directory_path_bloomberg + f"/final_{input_file.split('/')[-1]}" # "File path till last directory" + "/final_{Bloomberg_07-07-2023}.xlsx"
    # file_copy = str(input_file.split('/')[-2]) + f"/final_{input_file.split('/')[-1]}" # "Bloomberg" + "/final_{Bloomberg_07-07-2023}.xlsx"
    shutil.copyfile(input_file, file_copy)

    data = pd.read_excel(input_file)

    # st.write(input_file)
    
    # Extract the date and source/issuer from the filename
    last_part = input_file.split('/')[-1]
    date_str = last_part.split('_')[1].replace('.xlsx', '')
    source_issuer_str = last_part.split('_')[0]

    # st.write(source_issuer_str)
    
    new_data = pd.DataFrame({
        "ISIN": data["Security Code"],
        "Price": data["Indicative Valuation ( Mid )"],
        "Date": date_str,
        "Source": source_issuer_str, # Take the second part of "UBS/UBS_" or simply do "Source": new_prefix
        "Issuer": source_issuer_str # Take the second part "UBS/UBS_" or simply do "Source": new_prefix
    })

    output_file = file_copy
    
    # If you wish to create a seperate output file, use this command
    # output_file = input_file.replace('.csv', '_ISIN_updated.csv')
    new_data.to_excel(output_file, index=False)
        
    # print(f"Successfully created output file in the following path: {output_file}")
    
    return output_file