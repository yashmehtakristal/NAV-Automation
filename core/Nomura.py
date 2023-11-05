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

def rename_files_with_prefix_nomura(directory_path_nomura, prefix, new_prefix_nomura, extensions):
    """
    Rename files in the specified directory that start with a given prefix and have specified extensions.

    Args:
        directory_path (str): The path to the directory where the files are located.
        prefix (str): The prefix that files should start with to be considered for renaming.
        new_prefix (str): The new prefix to replace the old one.
        extensions (list): A list of file extensions to consider for renaming.

    Returns:
        None: file is updated in-place
    """
        
    files_in_directory = os.listdir(directory_path_nomura)
    
    # Tracking variables
    xls_files = 1
    pdf_files = 1

    for file_name in files_in_directory:

        if file_name.startswith(prefix) and any(file_name.lower().endswith(ext) for ext in extensions):

            if file_name.endswith(".xls"):
                new_file_name = f"{new_prefix_nomura}{xls_files}{os.path.splitext(file_name)[1]}"
                xls_files += 1
            
            else:
                new_file_name = f"{new_prefix_nomura}{pdf_files}{os.path.splitext(file_name)[1]}"
                pdf_files += 1

            old_file_path = os.path.join(directory_path_nomura, file_name).replace("\\","/")
            new_file_path = os.path.join(directory_path_nomura, new_file_name).replace("\\","/")

            os.rename(old_file_path, new_file_path)

            # print(f"Rename {file_name} to {new_file_name}")


def decrypt_and_save_excel(encrypted_file_path, decrypted_file_path, password):
    """
    Decrypts an encrypted Excel file and saves it as a decrypted xls file.

    Parameters:
        encrypted_file_path (str): The path to the encrypted Excel file.
        decrypted_file_path (str): The path to save the decrypted Excel file.
        password (str): The password used for encryption.

    Returns:
        None
    """
    
    try:
        with open(encrypted_file_path, 'rb') as encrypted_file:
            with open(decrypted_file_path, 'wb') as decrypted_file:
                
                office_file = msoffcrypto.OfficeFile(encrypted_file)
                office_file.load_key(password=password)
                office_file.decrypt(decrypted_file)
                
                # print(f"Successfully decrypted file in {decrypted_file_path}")
    except Exception as e:
        print(f"An error occurred while decrypting the file: {e}")


def process_excel_nomura(input_file_path, directory_path_nomura, new_prefix_nomura):
    """
    Process an Excel file to extract specific data and save it to another Excel file.

    Parameters:
    - input_file_path (str): The path to the input Excel file.
    - directory_path_nomura: Directory where all files are saved

    Returns:
    - output_file_path (str): The path where Excel file will be outputted
    - new_data (dataframe): The dataframe that we generated for that specific file
    """

    data = pd.read_excel(input_file_path)
    date_string = None

    for index, row in data.iterrows():
        if "Date" in row.values:
            date_index = list(row).index("Date")
            date_value = row[date_index + 1] # Extract the date value from the next column in the same row
            date_value = pd.to_datetime(date_value)
            date_string = date_value.strftime('%d-%m-%Y')
            break

    header_row_idx = None

    for idx, row in data.iterrows():
        if 'ISIN/Reference' in row.values:
            header_row_idx = idx
            break

    data = data.loc[header_row_idx:].reset_index(drop=True)
    data.columns = data.iloc[0]
    data = data.drop(0)
    
    if "Dirty Price" in data.columns:
        data = data[['ISIN/Reference', 'Dirty Price']].dropna()
    
    else:
        data = data[['ISIN/Reference', 'Indicative\nMid']].dropna()
        data = data.rename(columns={'Indicative\nMid': 'Indicative Mid'})

    
    # Create the new DataFrame based on updated requirements
    new_data = pd.DataFrame({
        "ISIN": data["ISIN/Reference"],
        # Essentially, the only difference between the 2 files from Nomura is this column difference between "Dirty Price" and "Indicative Mid"
        "Price": data["Dirty Price"] if "Dirty Price" in data.columns else data["Indicative Mid"],
        "Date": date_string,
        "Source": new_prefix_nomura, # Take the second part of "Nomura/Nomura_" or simply do "Source": new_prefix 
        "Issuer": new_prefix_nomura # Take the second part "Nomura/Nomura_" or simply do "Source": new_prefix
    })
    
    output_file_name = f"{new_prefix_nomura}_{date_string}.xlsx"
    output_file_path = os.path.join(directory_path_nomura, output_file_name).replace("\\","/")
        
    return output_file_path, new_data
