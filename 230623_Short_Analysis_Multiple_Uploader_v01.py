#!/usr/bin/env python
# coding: utf-8

# In[3]:


import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile


# In[2]:


name = st.text_input("File Name")
date = st.text_input("Today's date i.e. 231231")


## Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])


def combinesdds(xlsxs):
    sheet_names = ["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"]
    dfs = {name: [] for name in sheet_names}
    for xlsx in xlsxs:
        for sheet_name in sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_names)
            df = xlsx.parse(sheet_name)
            dfs[sheet_name].append(df)
    return dfs


def zipsdd_csvs(dfs,name,date):
    sheet_names = ["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"]
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
        dfs =  combinesdds(xlsxs)
        zip_io = zipsdd_csvs(dfs,name,date)
        tmp_download_link = st.download_button(
            label='Download CSVs',
            data=zip_io,
            file_name='dataframes.zip',
            mime='application/zip'
        )




# In[15]:





# In[17]:




