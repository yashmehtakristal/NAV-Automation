import streamlit as st
from streamlit_extras.app_logo import add_logo
from st_pages import Page, Section, add_page_title, show_pages, hide_pages

st.set_page_config(page_title="NAV Automation", page_icon="📖", layout="wide", initial_sidebar_state="expanded")
st.header("📖 NAV Automation")
add_logo("https://assets-global.website-files.com/614a9edd8139f5def3897a73/61960dbb839ce5fefe853138_Kristal%20Logotype%20Primary.svg")


# Display Markdown of the main page
st.markdown(
    '''
### Background

The primary purpose of this project is to facilitate the retrieval of Net Asset Value (NAV) data for structured notes.

Typically, this information is acquired directly from DBS, our custodian.

However, there are instances when DBS lacks reliable NAV sources. In such cases, we must obtain NAV values from various brokers.

Hence, the **objective** is to establish an automated process for this task and get the most updated NAV values as possible.

### How to run this program?

Simply head over to "NAV Automation" page. Upload a zip file which should be in the following format (case-sensitive!!):

 1. UBS
 2. Bloomberg
 3. Nomura
 4. Privatam

Each folder above should contain the files we typically get from each of the brokers. Place them in the respective folders, as is (don't change name or anything).

Here's an example of the final, zip file to be uploaded should look like:

 - **UBS**
	 - COB20230904---04Sep---2023-09-05-0644-20-CASH-362710-1.csv
	 - COB20230904---04Sep---2023-09-05-0644-20-CASH-362710-1.pdf
	 - COB20230904---04Sep---2023-09-05-0644-20-CASH-362710-1.xls
	 - disclaim.txt
 - **Bloomberg**
	 - Bloomberg_07-09-2023.xlsx
 - **Nomura**
	 - ValReq17729464.xls
	 - ValReq17729874.xls
	 - ValReq17729464.pdf
	 - ValReq17729874.pdf
 - **Privatam**
	 - Portfolio Report 2023-09-05.xlsx

You need not upload all files received from email, but it is recommended to do so as the program helps rename the files.

However, if you'd still want to upload particular files, please make sure the files below are minimally present (else program will not work):

 - **UBS**
	 - COB20230904---04Sep---2023-09-05-0644-20-CASH-362710-1.csv
 - **Bloomberg**
	 - Bloomberg_07-09-2023.xlsx
 - **Nomura**
	 - ValReq17729464.xls
	 - ValReq17729874.xls
 - **Privatam**
	 - Portfolio Report 2023-09-05.xlsx

At the end of the program, you will get 3 options: download final results excel file (.xlsx), download zip file, or individual excel files (.xlsx) of the various brokers. For most cases, you'd want to download the zip file. 

Finally, the ***desired final output*** you would want to look at is:
 1. **In root directory:** final_{Date when program was ran}.xlsx 
 2. **In respective subfolders:** "final_{Broker}_{Date when program was ran}"

### What to do if you get errors?

***Common solutions to particular errors*** include:

 1. Making sure zip file you upload is of above format (case-sensitive, right folders, right file extensions etc).
 2. Clear cache & cookies
 3. Run this on incognito mode (in case any extensions are interferring with this) 
 4. Check firewall settings

Apart from that, this code was created with ***certain assumptions*** in mind. These are, but not limited to the following:

 1. Broker provides consistent filename in that particular format
 2. Broker provides consistent filetype
 3. File content inside follows particular formatting (XYZ always present or at particular row etc)

Given how intricately dependent the code is on the input files, there may be errors appearing at some point in the future.

In that case, please contact yash.mehta@kristal.ai for assistance.
''')
