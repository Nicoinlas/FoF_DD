#!/usr/bin/env python
# coding: utf-8

# In[1]:

import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# In[2]:


name = st.text_input("Fund Short Name")
date = st.text_input("Quarter of Financials i.e. 231231")

uploaded_file = st.file_uploader("Choose file", type=["xlsx"])


def convertsdd_csv(xlsx, name, date):
    sheet_names = ["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"]
    zip_io = BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zipped_files:
        for sheet_name in sheet_names:
            df = xlsx.parse(sheet_name)
            file_name = str(date + " " + name + "_" + sheet_name + ".csv")
            csv_data = df.to_csv(index=False, encoding="utf-8-sig")
            zipped_files.writestr(file_name, csv_data.encode('utf-8-sig'))
    zip_io.seek(0)
    return zip_io

if uploaded_file is not None and name is not None and date is not None:
    xlsx = pd.ExcelFile(uploaded_file)
    zip_io = convertsdd_csv(xlsx, name, date)
    
    up1 = xlsx.parse("UP01_Funds")
    up2 = xlsx.parse("UP02_Fund Financials")
    up3 = xlsx.parse("UP03_Portfolio Companies")
    up4 = xlsx.parse("UP04_PFC Financials")

    st.write("Funds")
    st.dataframe(up1)
    st.write("Fund Financials")
    st.dataframe(up2)
    st.write("PFCs")
    st.dataframe(up3)
    st.write("PFC Financials")
    st.dataframe(up4)

    btn = st.button("Download CSVs")

    if btn:
        tmp_download_link = st.download_button(
            label='Download CSVs',
            data=zip_io,
            file_name='dataframes.zip',
            mime='application/zip'
        )




# In[4]:




