import msoffcrypto
import xlrd
import openpyxl
import os
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pickle
import streamlit as st


### UBS

def rename_files_with_prefix_ubs(directory_path, prefix, new_prefix, extensions):
    
    """
    Rename files in the specified directory that start with a given prefix and have specific extensions.

    Args:
        directory_path (str): The path to the directory where the files are located.
        prefix (str): The prefix that files should start with to be considered for renaming.
        new_prefix (str): The new prefix to include in filename
        extensions (list): List of extensions you need to parse

    Returns:
        csv_file_path_ubs (list): List of the csv files (main data source) of UBS
    """

    csv_file_path_ubs = []
    files_in_directory = os.listdir(directory_path)

    for file_name in files_in_directory:

        if file_name.startswith(prefix) and any(file_name.lower().endswith(ext) for ext in extensions):

            date_part = file_name[len(prefix):len(prefix) + 8]
            formatted_date = f"{date_part[6:]}-{date_part[4:6]}-{date_part[:4]}"
            new_file_name = f"{new_prefix}_{formatted_date}{os.path.splitext(file_name)[1]}"
            old_file_path = os.path.join(directory_path, file_name).replace("\\","/")
            new_file_path = os.path.join(directory_path, new_file_name).replace("\\","/")
                        
            if new_file_path.endswith('.csv'):
                csv_file_path_ubs.append(new_file_path)
                                        
            os.rename(old_file_path, new_file_path)
                                          
            # print(f"Successfully renamed {file_name} to {new_file_name}")
            
    return csv_file_path_ubs


def process_csv_ubs(directory_path, input_file):
    
    """
    Processes the input CSV file and updates it according to desired output specification.

    Parameters:
        input_file (str): The path to the input CSV file that needs to be processed

    Returns:
        output_file (str) - The path to the output xlsx file that needs to be processed
    """

    data = pd.read_csv(input_file)
    # st.write(input_file)
    date_str = input_file.split('_')[1].replace('.csv', '')
    source_issuer_str = input_file.split('_')[0]
    
    last_part = input_file.split('/')[-1]
    new_last_part = last_part.replace('.csv','.xlsx')

    # st.write(source_issuer_str)
    
    new_data = pd.DataFrame({
        "ISIN": data["Security Code"],
        "Price": data["Indicative Valuation ( Mid )"],
        "Date": date_str,
        "Source": source_issuer_str.split('/')[-2], # Take the second part of "UBS/UBS_" or simply do "Source": new_prefix
        "Issuer": source_issuer_str.split('/')[-2] # Take the second part "UBS/UBS_" or simply do "Source": new_prefix
    })
    
    output_file_path = directory_path + f"/final_{new_last_part}"

    # st.write(output_file_path)
        
    new_data.to_excel(output_file_path, index=False)

    # print(f"Successfully created output file in {output_file}")
    
    return output_file_path
