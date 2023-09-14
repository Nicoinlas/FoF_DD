#!/usr/bin/env python
# coding: utf-8

# In[3]:


import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import numpy as np

# Define your options
options = ['Short Analysis', 'PQ Financials', 'PQ Fund', 'PQ Fund Manager','Long Analysis']
selected_option = st.selectbox('Select your options:', options)
# Display the selected options
st.write('You selected:', selected_option)

name = st.text_input("File Name")
date = st.text_input("Date will be shown on file i.e. 231231")


## Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])


arrays = {
    'Short Analysis':["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"],
    'PQ Financials': ["08a_UP SF_Financial_New","08b_UP SF_Financial_New","09a_UP SF_Financial_Existing","09b_UP SF_Financial_Existing"],
    'PQ Fund': ["05 Upload SF_Fund_ Existing_A","06 Upload SF_Fund_ Existing_B","07 Upload SF_Fund_New_A","07 Upload SF_Fund_New_B"],
    'PQ Fund Manager':["04_SF Upload_New","05_SF Upload_Existing"],
    'Long Analysis':["UP_01 Fund Manager","UP_02A Funds","UP_02B Funds", "UP_03 Fund Financials", "UP_04A PFCs","UP_04B PFCs", "UP_05 PFC Financials","UP_06 Team", "UP_08 Deal Attribution","UP_09 IC Memo"],
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
            dfs["UP01_Funds"] = dfs["UP01_Funds"].sort_values(['Salesforce.com ID', 'Date of Latest Quarterly Performance'], ascending=[True,False], na_position='last')
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
                df[col] = pd.to_datetime(df[col], format='mixed')
                
                # Format the datetime column to the desired format
                df[col] = df[col].dt.strftime('%Y-%m-%dT%H:%M:%SZ')
                      
            if "Vintage PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Vintage PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Vintage PQ"] = df["Vintage PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)
            
            if "Vintage" in df.columns:
                # Fill NaNs with a dummy value
                df["Vintage"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Vintage"] = df["Vintage"].apply(lambda x: str(int(x)) if x != -1 else np.nan)
           
            if "Year established PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Year established PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Year established PQ"] = df["Year established PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)
                 
            if "Fund ID PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Fund ID PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Fund ID PQ"] = df["Fund ID PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)

            if "Company ID PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Company ID PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Company ID PQ"] = df["Company ID PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)

            if "Fund ID Sub Strategy PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Fund ID Sub Strategy PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Fund ID Sub Strategy PQ"] = df["Fund ID Sub Strategy PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)

            if "Quartile Rank PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Quartile Rank PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Quartile Rank PQ"] = df["Quartile Rank PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)

            if "Preqin Quartile Rank PQ" in df.columns:
                # Fill NaNs with a dummy value
                df["Preqin Quartile Rank PQ"].fillna(-1, inplace=True)
                
                # Convert to integer, then to string, and replace the dummy value with NaN
                df["Preqin Quartile Rank PQ"] = df["Preqin Quartile Rank PQ"].apply(lambda x: str(int(x)) if x != -1 else np.nan)
            
                         


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


