import streamlit as st 
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
import re
import zipfile
from collections import Counter
import io
import base64  # Import the base64 module

def create_pickle_file(new_directory_path, current_date):

    # pickle_directory = f"{new_directory_path}/Pickle"
    pickle_directory = f"{new_directory_path}"

    # If Pickle folder does not exist, make folder
    if not os.path.exists(pickle_directory):
        os.makedirs(pickle_directory)

    # Create pickle file saving the current date of iteration
    with open(f'{pickle_directory}/last_run_date.pkl', 'wb') as file:
        pickle.dump(current_date, file)


def load_pickle_file(directory):

    with open(f'{directory}/last_run_date.pkl', 'rb') as file:
        last_run_date = pickle.load(file)  

    return last_run_date


def seeing_most_common_date(output_files):
    dates = [re.search(r'\d{2}-\d{2}-\d{4}', name).group() for name in output_files]
    date_counts = Counter(dates)
    most_common_date = date_counts.most_common(1)[0][0]
    date_parts = most_common_date.split('-')
    formatted_most_common_date = f"{date_parts[0]}-{date_parts[1]}-{date_parts[2]}"


def creating_broker_folders():

    # List of folder names
    folder_names = ['UBS', 'Privatam', 'Nomura', 'Bloomberg', 'Catfp']

    # Get the current directory
    current_directory = os.getcwd()

    # Iterate over the folder names and create them in the current directory
    for folder_name in folder_names:
        folder_path = os.path.join(current_directory, folder_name)
        
        # Creating folder path if it does not exist
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            # print(f'Created folder: {folder_name} in {current_directory}')
        
        else:
            pass
            # print(f'Folder {folder_name} already exists in {current_directory}')


def create_historical_folders(current_directory):

    # Get the current directory
    # current_directory = os.getcwd()

    # Create "Historical" folder in the current directory if it does not exist
    historical_folder_path = os.path.join(current_directory, 'Historical')

    if not os.path.exists(historical_folder_path):

        os.makedirs(historical_folder_path)
        # print(f'Created "Historical" folder in {current_directory}')
    
    else:
        pass
        # print(f'"Historical" folder already exists in {current_directory}')


    # Iterate over all items (folders and files) in the current directory
    for item in os.listdir(current_directory):
        item_path = os.path.join(current_directory, item)
        
        # Check if the item is a directory (folder) and does not have certain names like "Historical" & ".ipynb_checkpoints"
        if os.path.isdir(item_path) and item not in ('Historical', '.ipynb_checkpoints', '.git'):
            historical_folder_path = os.path.join(item_path, 'Historical')
            
            # Check if the "Historical" folder already exists, and create it if not
            if not os.path.exists(historical_folder_path):
                os.makedirs(historical_folder_path)
                # print(f'Created "Historical" folder in {item}')
            else:
                pass
                # print(f'"Historical" folder already exists in {item}')


def move_files_into_historical_folder(directory, last_run_date):
    '''
    Move files in the specified directory into the "Historical" folder for a particular directory & its sub-directories
    '''

    # with open(f'{directory}/last_run_date.pkl', 'rb') as file:
    #     last_run_date = pickle.load(file)

    # Create the Historical folder and last_run_date folder in the specified directory level
    historical_root_folder = os.path.join(directory, "Historical")
    last_run_date_root_folder = os.path.join(historical_root_folder, last_run_date)
        
    # Make the last_run_date folder in the historical_root_folder
    if not os.path.exists(last_run_date_root_folder):
        os.makedirs(last_run_date_root_folder)
    
    # Move final_xlsx and last_run_date.pkl to the last_run_date folder at root level
    
    final_xlsx_source = os.path.join(directory, f"final_{last_run_date}.xlsx")
    final_xlsx_target = os.path.join(last_run_date_root_folder, f"final_{last_run_date}.xlsx")
    shutil.move(final_xlsx_source, final_xlsx_target)
    
    pickle_source = os.path.join(directory, "last_run_date.pkl")
    pickle_target = os.path.join(last_run_date_root_folder, "last_run_date.pkl")
    shutil.move(pickle_source, pickle_target)
    
    # For each broker folder, create Historical and last_run_date folders, then move files
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        
        # Check if the item is a directory (folder) and is not the root Historical or other excluded folders
        if os.path.isdir(item_path) and item not in ('Historical', '.ipynb_checkpoints'):
            
            # Construct path to historical subfolder inside broker folder
            historical_broker_folder = os.path.join(item_path, "Historical")
            
            # Construct path to last run date subfolder inside historical subfolder
            last_run_date_broker_folder = os.path.join(historical_broker_folder, last_run_date)
            
            # Create last run date subfolder inside historical subfolder if it does not exist
            if not os.path.exists(last_run_date_broker_folder):
                os.makedirs(last_run_date_broker_folder)
            
            # Move files from broker folder to the last_run_date_broker_folder
            for file in os.listdir(item_path):
                file_path = os.path.join(item_path, file)
                if os.path.isfile(file_path):
                    shutil.move(file_path, last_run_date_broker_folder)


# Call function
# move_files_into_historical_folder(".")


def upload_zip_files():

    # Create a temporary directory, called zipfiles, for storing uploaded ZIP files
    # temp_dir = "zipfiles"
    # os.makedirs(temp_dir, exist_ok=True)

    # Upload ZIP files
    uploaded_zip_file = st.file_uploader(
        'Upload a ZIP file',
        type=["zip"],
        accept_multiple_files = False,
        key = "zip_uploader",
        help = "Please upload zip file according to particular order specified in about section",
        disabled = False,
        label_visibility = "visible"
        )

    return uploaded_zip_file
    
def create_zipfile_directory():

    # Check if the "zipfiles" folder exists, and if not, create it
    directory_path = "zipfiles/"
    new_directory_name = None
    new_directory_path = None

    # Check if the "zipfiles" folder (master folder) exists, and if not, create it
    if not os.path.exists(directory_path):
        os.makedirs(directory_path)

    # If there are no folders in "zipfiles" folder (master folder), create a new default folder, "iteration0"
    if not os.listdir(directory_path):
        new_directory_name = "iteration0"

    # List all directories in the "embeddings" directory
    directories = os.listdir(directory_path)

    # Filter directories that start with "iteration" (8 characters) and are followed by digits
    numeric_directories = [dirname for dirname in directories if dirname.startswith("iteration") and dirname[9:].isdigit()]

    # If there is a list like that
    if numeric_directories:

        # Sort the numeric directory names in descending order
        sorted_directories = sorted(numeric_directories, key=lambda x: int(x[9:]), reverse=True)
        
        # The first element in the sorted list will be the largest directory name
        largest_directory = sorted_directories[0]
        
        # Extract the numeric part of the largest directory name and increment it by 1
        largest_number = int(largest_directory[9:])

        # This is finding the new_number for the new directory name
        new_number = largest_number + 1
        
        # Create the new directory name with the incremented number
        new_directory_name = f"iteration{new_number}"

    # If new_directory_name is set, create the new directory
    if new_directory_name:

        new_directory_path = os.path.join(directory_path, new_directory_name)
        os.makedirs(new_directory_path)

    # Write info of Directory path and Directory name - Diagnostic

    # st.write(f"Master folder where all directories are stored: {directory_path}")
    # st.write(f"Directory path where result will be stored: {new_directory_path}") 
    # st.write(f"New Directory (Folder) name: {new_directory_name}")

    # st.write(f"Master folder where all directories are stored: {directory_path}") Example: zipfiles/
    # st.write(f"Directory path where result will be stored: {new_directory_path}") Example: zipfiles/iteration0
    # st.write(f"New Directory (Folder) name: {new_directory_name}") Example: iteration0

    return directory_path, new_directory_path, new_directory_name


def write_zip_files_to_directory(uploaded_zip_file, new_directory_path):

    # Determine the full path to save the ZIP file
    zip_file_path = os.path.join(new_directory_path, uploaded_zip_file.name)

    # Save the uploaded ZIP file to the temporary directory
    with open(zip_file_path, "wb") as f:
        f.write(uploaded_zip_file.read())

    # Unzip the uploaded ZIP file
    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(new_directory_path)

    # Delete the uploaded ZIP file after extraction
    os.remove(zip_file_path)

    st.success(f"ZIP file '{uploaded_zip_file.name}' uploaded and extracted in: {new_directory_path}", icon="✅")


def create_zip_link(directory, zip_filename):
    # st.write(f"Creating a zip file of all files in this directory and its subdirectories: {directory}")

    zip_buffer = io.BytesIO()

    # shutil.make_archive(zip_buffer, 'zip', directory)

    # Create a zip archive in memory
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:

        # foldername: Path to current folder being processed
        # subfolders: List of subdirectory names within current folder
        # filenames: List of file names within current folder

        # Note: OS.walk skips over empty folders

        for foldername, subfolders, filenames in os.walk(directory, followlinks = True):

            zipf.write(foldername, os.path.relpath(foldername, directory))

            for filename in filenames:

                # Create the full file path
                file_path = os.path.join(foldername, filename)

                # How would you refer to the file from within the directory (without including the full path)
                # Maintains structure 
                arcname = os.path.relpath(file_path, directory)

                # Arcname specifies name to be used for file within ZIP archive
                zipf.write(file_path, arcname)

    # st.write("Zip file created successfully")

    # Seek back to the beginning of the BytesIO object
    zip_buffer.seek(0)

    # Generate the data URI for the ZIP file
    data_uri = f"data:application/zip;base64,{base64.b64encode(zip_buffer.read()).decode()}"

    # Create download link
    download_link = f"<a href='{data_uri}' download='{zip_filename}.zip'>Download zip file</a>"
    st.markdown(download_link, unsafe_allow_html=True) # Display the custom download link


    # Download zip file via button
    # For zip file, use MIME type = “application/zip”
    # For rar and 7z file, use MIME type = “application/octet-stream”
    # zip_file_download = st.download_button(
    #         label = "Download zip file",
    #         data = zip_buffer.read(),
    #         file_name = f"{zip_filename}.zip",
    #         mime = "application/zip"
    #     )

def create_xlsx_link(xlsx_path, filename, filename_last_broker):

    with open(xlsx_path, "rb") as file:
        excel_bytes = file.read()
        excel_b64 = base64.b64encode(excel_bytes).decode()

    st.markdown(f'<a href="data:application/vnd.ms-excel;base64,{excel_b64}" download="{filename}.xlsx">Download {filename_last_broker} Excel file</a>', unsafe_allow_html=True)

def create_xlsx_link_after_moving_historical(xlsx_path, filename, filename_last_broker):

    with open(xlsx_path, "rb") as file:
        excel_bytes = file.read()
        excel_b64 = base64.b64encode(excel_bytes).decode()

    st.markdown(f'<a href="data:application/vnd.ms-excel;base64,{excel_b64}" download="{filename}.xlsx">Download {filename_last_broker} Excel file</a>', unsafe_allow_html=True)

def download_dataframe_as_excel_link(data, date):

    # buffer to use for excel writer
    buffer = io.BytesIO()

    downloaded_file = data.to_excel(buffer, index=False, header=True)

    buffer.seek(0)  # reset pointer

    b64 = base64.b64encode(buffer.read()).decode()  # some strings

    link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="final_{date}.xlsx">Download final results excel file</a>'

    st.markdown(link, unsafe_allow_html=True)


def create_pdf_link(pdf_path, filename, filename_last_broker):

    with open(pdf_path, "rb") as file:
        pdf_bytes = file.read()
        pdf_b64 = base64.b64encode(pdf_bytes).decode()

    st.markdown(f'<a href="data:application/pdf;base64,{pdf_b64}" download="{filename}.pdf">Download {filename_last_broker} PDF</a>', unsafe_allow_html=True)

