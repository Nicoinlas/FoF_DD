#!/usr/bin/env python
# coding: utf-8

# In[3]:


import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import numpy as np

# Define your options
options = ['Short Analysis', 'PQ Financials', 'PQ Fund', 'PQ Fund Manager']
selected_option = st.selectbox('Select your options:', options)
# Display the selected options
st.write('You selected:', selected_option)

name = st.text_input("File Name")
date = st.text_input("Date shown on file i.e. 231231")


## Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])


arrays = {
    'Short Analysis':["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"],
    'PQ Financials': ["08a_UP SF_Financial_New","08b_UP SF_Financial_New","09a_UP SF_Financial_Existing","09b_UP SF_Financial_Existing"],
    'PQ Fund': ["05 Upload SF_Fund_ Existing_A","06 Upload SF_Fund_ Existing_B","07 Upload SF_Fund_New_A","07 Upload SF_Fund_New_B"],
    'PQ Fund Manager':["04_SF Upload_New","05_SF Upload_Existing"],
}


sheet_names = arrays[selected_option]

def combine(xlsxs, sheet_names):
    dfs = {name: [] for name in sheet_names}
    for xlsx in xlsxs:
        for sheet_name in sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            # Remove empty rows
            df = df.replace(' ', np.nan)
            df = df.dropna(how='all')
            dfs[sheet_name].append(df)
    for sheet_name in sheet_names:
        dfs[sheet_name] = pd.concat(dfs[sheet_name], ignore_index=True)
    if selected_option == 'Short Analysis':
        if 'Salesforce.com ID' in dfs["UP01_Funds"].columns and 'Date of Latest Quarterly Performance' in dfs["UP01_Funds"].columns:
            dfs["UP01_Funds"] = dfs["UP01_Funds"].sort_values(['Salesforce.com ID', 'Date of Latest Quarterly Performance'], ascending=[True,False])
        if 'Salesforce.com ID' in dfs["UP01_Funds"].columns:
            dfs["UP01_Funds"] = dfs["UP01_Funds"].drop_duplicates(subset='Salesforce.com ID', keep='first')
    return dfs


import pandas as pd

import pandas as pd
import numpy as np

def zipsdd_csvs(dfs,name,date, sheet_names):
    zip_io = BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zipped_files:
        for sheet_name in sheet_names:
            # Get the dataframe
            df = dfs[sheet_name]
            
            # Find all columns that contain 'date' in the column name
            date_cols = [col for col in df.columns if 'date' in col.lower()]
            
            # Convert and format datetime columns
            for col in date_cols:
                # Ensure the column is in datetime format
                df[col] = pd.to_datetime(df[col])
                
                # Format the datetime column to the desired format
                df[col] = df[col].dt.strftime('%Y-%m-%dT%H:%M:%SZ')
            
            # Find all numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            # Convert numbers with 0 decimal to integer
            for col in numeric_cols:
                df[col] = df[col].apply(lambda x: int(x) if x == x // 1 else x)
            
            if "Vintage PQ" in df.columns:
                df["Vintage PQ"] = """ + df["Vintage PQ"].astype(str) + """

            file_name = str(date + " " + name + "_" + sheet_name + ".csv")
            csv_data = df.to_csv(index=False, encoding="utf-8-sig")
            zipped_files.writestr(file_name, csv_data.encode('utf-8-sig'))
    zip_io.seek(0)
    return zip_io


btn = st.button("Download CSVs")

if btn:
    dfs =  combine(xlsxs, sheet_names)
    zip_io = zipsdd_csvs(dfs,name,date, sheet_names)
    tmp_download_link = st.download_button(
        label='Download CSVs',
        data=zip_io,
        file_name='dataframes.zip',
        mime='application/zip'
        )


