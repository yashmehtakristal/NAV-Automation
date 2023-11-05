import msoffcrypto
import xlrd
import openpyxl
import os
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pickle


def create_final_results_file(output_files, current_date, new_directory_path):
    
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

    combined_data = pd.DataFrame()

    for file_path in output_files:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            print(f"Unsupported file format: {file_path}")
            continue

        combined_data = pd.concat([combined_data, df], ignore_index=True)

    results_output_file_path = f'{new_directory_path}/final_{current_date}.xlsx'
    workbook = Workbook()
    worksheet = workbook.active

    for row in dataframe_to_rows(combined_data, index=False, header=True):
        worksheet.append(row)
        
    workbook.save(results_output_file_path)

    return results_output_file_path, combined_data
