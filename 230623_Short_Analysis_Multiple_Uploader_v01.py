import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

name = st.text_input("File Name")
date = st.text_input("Today's date i.e. 231231")


## Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])


def combinesdds(xlsxs):
    sheet_names = ["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"]
    dfs = {name: [] for name in sheet_names}
    for xlsx in xlsxs:
        for sheet_name in sheet_names:
            df = pd.read_excel(xlsx, sheet_name=sheet_name)
            dfs[sheet_name].append(df)
    for sheet_name in sheet_names:
        dfs[sheet_name] = pd.concat(dfs[sheet_name], ignore_index=True)
    if 'Salesforce.com ID' in dfs["UP01_Funds"].columns and 'Date of Latest Quarterly Performance' in dfs["UP01_Funds"].columns:
        dfs["UP01_Funds"] = dfs["UP01_Funds"].sort_values(['Salesforce.com ID', 'Date of Latest Quarterly Performance'], ascending=[True,False])
    if 'Salesforce.com ID' in dfs["UP01_Funds"].columns:
        dfs["UP01_Funds"] = dfs["UP01_Funds"].drop_duplicates(subset='Salesforce.com ID', keep='first')
    return dfs


def zipsdd_csvs(dfs,name,date):
    sheet_names = ["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"]
    zip_io = BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zipped_files:
        for sheet_name in sheet_names:
            file_name = str(date + " " + name + "_" + sheet_name + ".csv")
            csv_data = dfs.to_csv(index=False, encoding="utf-8-sig")
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


