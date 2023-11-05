import streamlit as st
from streamlit_extras.app_logo import add_logo
from st_pages import Page, Section, add_page_title, show_pages, hide_pages

st.set_page_config(page_title="NAV Automation", page_icon="📖", layout="wide", initial_sidebar_state="expanded")
st.header("📖 NAV Automation")
add_logo("https://assets-global.website-files.com/614a9edd8139f5def3897a73/61960dbb839ce5fefe853138_Kristal%20Logotype%20Primary.svg")
show_pages(
    [
        Page("pages/about.py", "About", "😀"),
        Page("upload_zip_basic.py","NAV Automation", "🗂️"),
        # Page("pages/config.py", "Config", "⚙️"),
        # Page("pages/upload_individual_advanced.py", "Upload individual (advanced)", "📂")
    ]
)

import msoffcrypto
import xlrd
import openpyxl
import os
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pickle

from core.UBS import rename_files_with_prefix_ubs, process_csv_ubs
from core.Bloomberg import process_excel_bloomberg, getting_file_paths_bloomberg
from core.Nomura import rename_files_with_prefix_nomura, decrypt_and_save_excel, process_excel_nomura
from core.Privatam import rename_files_with_prefix, keep_only_products_worksheet, process_excel_privatam
from core.Final import create_final_results_file
from other.helper_funcs import create_pickle_file, upload_zip_files, create_zipfile_directory, write_zip_files_to_directory, create_zip_link, create_zip_link, create_xlsx_link, create_historical_folders, move_files_into_historical_folder, load_pickle_file, download_dataframe_as_excel_link


# initialize global variables
# output_files = []

# st.header("NAV Automation")


uploaded_zip_file = upload_zip_files()

if uploaded_zip_file is not None:

    if st.button("Start NAV Automation", type = "primary"):
            
        output_files = []

        with st.spinner("Extract zip files"):
            directory_path, new_directory_path, new_directory_name = create_zipfile_directory()
            write_zip_files_to_directory(uploaded_zip_file = uploaded_zip_file, new_directory_path = new_directory_path)
            upload_zip_file_name = uploaded_zip_file.name
            final_upload_zip_file_name = upload_zip_file_name.split('.')[0]
            # st.write(final_upload_zip_file_name)

        ### UBS
        with st.spinner(text = "Extracting NAVs from UBS documents"):

            # Initializing variables for UBS
            directory_path_ubs = f"{new_directory_path}/UBS"
            prefix = 'COB'
            new_prefix_ubs = 'UBS'
            extensions = ['.csv', '.xlsx', '.xls', '.pdf']

            excel_file_path_ubs = rename_files_with_prefix_ubs(directory_path = directory_path_ubs, prefix = prefix, new_prefix = new_prefix_ubs, extensions = extensions)

            for excel_file_path in excel_file_path_ubs:
                output_file = process_csv_ubs(directory_path = directory_path_ubs, input_file = excel_file_path)
                output_files.append(output_file)

        st.success("Successfully extracted NAVs from UBS documents", icon="✅")


        ### Bloomberg
        with st.spinner(text = "Extracting NAVs from Bloomberg documents"):

            # Initializing variables for Bloomberg
            directory_path_bloomberg = f"{new_directory_path}/Bloomberg"
            xls_file_path_bloomberg = []

            xls_file_path_bloomberg = getting_file_paths_bloomberg(directory_path_bloomberg = directory_path_bloomberg, xls_file_path_bloomberg = xls_file_path_bloomberg)

            # st.write(xls_file_path_bloomberg)

            for xls_file_path in xls_file_path_bloomberg:
                # print(xls_file_path)
                output_file = process_excel_bloomberg(directory_path_bloomberg = directory_path_bloomberg, input_file = xls_file_path)
                output_files.append(output_file)

        st.success("Successfully extracted NAVs from Bloomberg documents", icon="✅")


        ### Nomura
        with st.spinner(text = "Extracting NAVs from Nomura documents"):

            # Initializing variables
            directory_path_nomura = f"{new_directory_path}/Nomura"
            prefix = 'ValReq'
            new_prefix_nomura = 'Nomura'
            extensions = ['.csv', '.xlsx', '.xls', '.pdf']
            password = '11400981'

            # Calling functions
            rename_files_with_prefix_nomura(directory_path_nomura = directory_path_nomura, prefix = prefix, new_prefix_nomura = new_prefix_nomura, extensions = extensions)

            encrypted_file_path = f'{directory_path_nomura}/Nomura1.xls'
            decrypted_file_path = f'{directory_path_nomura}/Nomura1-decrypted.xls'
            encrypted_file_path2 = f'{directory_path_nomura}/Nomura2.xls'
            decrypted_file_path2 = f'{directory_path_nomura}/Nomura2-decrypted.xls'

            decrypt_and_save_excel(encrypted_file_path = encrypted_file_path, decrypted_file_path = decrypted_file_path, password = password)
            decrypt_and_save_excel(encrypted_file_path = encrypted_file_path2, decrypted_file_path = decrypted_file_path2, password = password)

            file_1_path, file_1_data = process_excel_nomura(input_file_path = decrypted_file_path, directory_path_nomura = directory_path_nomura, new_prefix_nomura = new_prefix_nomura)
            file_2_path, file_2_data = process_excel_nomura(input_file_path = decrypted_file_path2, directory_path_nomura = directory_path_nomura, new_prefix_nomura = new_prefix_nomura)

            last_part = file_1_path.split('/')[-1]
            # new_last_part = last_part.replace('.xlsx','')

            new_file_path = directory_path_nomura + f'/final_{last_part}'

            final_df = pd.concat([file_1_data, file_2_data], axis=0)
            excel_file = new_file_path # Can also use file_2_path as both are same
            final_df.to_excel(excel_file, index=False)
            output_files.append(new_file_path)


        st.success("Successfully extracted NAVs from Nomura documents", icon="✅")


        ### Privatam
        with st.spinner(text = "Extracting NAVs from Privatam documents"):

            # Initializing Privatam variables
            directory_path_privatam = f"{new_directory_path}/Privatam" 
            prefix = 'Portfolio Report '
            new_prefix = 'Privatam'
            extensions = ['.xlsx']

            # Calling Privatam functions
            excel_file_path_privatam = rename_files_with_prefix(directory_path = directory_path_privatam, prefix = prefix, new_prefix = new_prefix, extensions = extensions)

            for excel_file_path in excel_file_path_privatam:
                keep_products = keep_only_products_worksheet(excel_file_path)
                output_file = process_excel_privatam(directory_path_privatam = directory_path_privatam, input_file_path = excel_file_path)
                output_files.append(output_file)


        ### Privatam
        with st.spinner(text = "Creating final results file"):
            
            ### Creating final results file
            current_date = datetime.now().strftime('%d-%m-%Y')
            results_output_file_path, final_results_dataframe = create_final_results_file(output_files, current_date, new_directory_path)

        st.success("Successfully created final results file", icon="✅")


        with st.spinner(text = "Conducting Post-Processing"):

            # Creating pickle file
            create_pickle_file(new_directory_path = new_directory_path, current_date = current_date)

        st.success("Successfully finished post-processing", icon="✅")

        st.markdown("### Download final results")

        st.dataframe(data = final_results_dataframe, use_container_width = True, column_order = None)

        download_dataframe_as_excel_link(data = final_results_dataframe, date = current_date)

        st.markdown("### Download zip file")

        create_zip_link(directory = new_directory_path, zip_filename = final_upload_zip_file_name)

        st.markdown("### Download individual files")

        for i in range(len(output_files)):
            filename_last = output_files[i].split('/')[-1]
            filename_last_broker = filename_last.split('_')[1]
            filename_last_filter = filename_last.split('.')[0]
            create_xlsx_link(xlsx_path = output_files[i], filename = filename_last_filter, filename_last_broker = filename_last_broker)

else:
    st.warning("Please upload a zip file which will be used in the NAV automation process")

