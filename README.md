# ISIN Processor Automation - Kristal.AI

## Dependencies needed to install

It is strongly recommended to establish a pip environment for executing all the files.

The code was generated using python 3.11.4 and any python version 3.7+ should be consistent with executing our code.

|Sr No.|Library used|Purpose| Installation|
|--|--| --| --|
|1| io| Used for dealing with for dealing with various types of I/O| Built-in python|
|2| msoffcrypto| Used for decrypting encrypted MS Office files with password| pip install msoffcrypto-tool|
|3| xlrd| Used for reading data and formatting information from Excel files in `.xls`  format.| pip install xlrd|
|4| openpyxl| Used for reading/writing Excel 2010 xlsx/xlsm files.| pip install openpyxl |
|5| os| Used to access operating system dependent functionality in a portable manner| Built-in python| 
|6| pandas| Used for data manipulation and analysis, offering powerful data structures and tools for working with structured data.| pip install pandas|
|7| datetime| Used for manipulating dates and times| Built-in python| 

## Instructions to run code

Please follow the below instructions for setting our code:

 Utilize <b> git bash (Windows) or powershell or Mac/Linux:</b>
 
Step 1: ```git clone https://github.com/yashmehtaym/ISINProcessor.git```

Step 2: ``` git checkout main```

## Input/Output specifications

Setup folders for the different sources/brokers and insert your initial data here like:

 1. UBS
 2. Citi
 3. Nomura
 4. Privatam
 5. Catfp

Once, you run the python file, inside each of the specific folders, you will see output file generated in the default format as such:

(Broker)_(Date).xlsx

This contains the ISIN format generated for that specific broker in the broker's folder.


Additionally, you will see another file generated like:

Final.xlsx

This contains the combined output from all sources/brokers and is the final file, you should take a look at.
