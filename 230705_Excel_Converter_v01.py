#!/usr/bin/env python
# coding: utf-8

# In[3]:


import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile


# In[18]:


name = st.text_input("File Name")
date = st.text_input("Today's date i.e. 231231")


## Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])

# Define your options
options = ['Short Analysis', 'PQ Financials', 'PQ Fund', 'PQ Fund Manager']
selected_option = st.selectbox('Select your options:', options)
# Display the selected options
st.write('You selected:', selected_option)

arrays = {
    'Short Analysis':["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"],
    'PQ Financials': ["08a_UP SF_Financial_New","08b_UP SF_Financial_New","09a_UP SF_Financial_Existing","09b_UP SF_Financial_Existing"],
    'PQ Fund': [],
    'PQ Fund Manager':[],
}


sheet_names = arrays[selected_option]

def combinesdds(xlsxs, sheet_names):
    dfs = {name: [] for name in sheet_names}
    for xlsx in xlsxs:
        for sheet_name in sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
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


def zipsdd_csvs(dfs,name,date, sheet_names):
    zip_io = BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zipped_files:
        for sheet_name in sheet_names:
            file_name = str(date + " " + name + "_" + sheet_name + ".csv")
            csv_data = dfs[sheet_name].to_csv(index=False, encoding="utf-8-sig")
            zipped_files.writestr(file_name, csv_data.encode('utf-8-sig'))
    zip_io.seek(0)
    return zip_io

btn = st.button("Download CSVs")

if btn:
    dfs =  combinesdds(xlsxs, sheet_names)
    zip_io = zipsdd_csvs(dfs,name,date, sheet_names)
    tmp_download_link = st.download_button(
        label='Download CSVs',
        data=zip_io,
        file_name='dataframes.zip',
        mime='application/zip'
        )



