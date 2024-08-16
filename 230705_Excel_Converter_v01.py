import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import numpy as np
import datetime
import logging

# Set up logging
logging.basicConfig(filename='conversion_errors.log', level=logging.ERROR, 
                    format='%(asctime)s %(message)s')

# Define your options
options = ['Short Analysis', 'PQ Financials', 'PQ Fund', 'PQ Fund Manager','Long Analysis']
selected_option = st.selectbox('Select your options:', options)
# Display the selected options
st.write('You selected:', selected_option)

name = st.text_input("File Name")

# Get the current date, Format the date as YYMMDD
today = datetime.date.today()
date = str(today.strftime('%y%mdd'))

# Upload multiple files
xlsxs = st.file_uploader("Upload your files here...", accept_multiple_files=True, type=["xlsx"])

arrays = {
    'Short Analysis':["UP01_Funds", "UP02_Fund Financials", "UP03_Portfolio Companies", "UP04_PFC Financials"],
    'PQ Financials': ["08a_UP SF_Financial_New","08b_UP SF_Financial_New","09a_UP SF_Financial_Existing","09b_UP SF_Financial_Existing"],
    'PQ Fund': ["05 Upload SF_Fund_ Existing_A","06 Upload SF_Fund_ Existing_B","07 Upload SF_Fund_New_A","07 Upload SF_Fund_New_B"],
    'PQ Fund Manager':["04_SF Upload_New","05_SF Upload_Existing"],
    'Long Analysis':["UP_01 Fund Manager","UP_02A Funds","UP_02B Funds", "UP_03 Fund Financials", "UP_04A PFCs","UP_04B PFCs", "UP_05 PFC Financials","UP_06 Team", "UP_07 Deal Attribution","UP_08 IC Memo"],
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

def zipsdd_csvs(dfs, name, date, sheet_names):
    zip_io = BytesIO()
    with zipfile.ZipFile(zip_io, mode='w') as zipped_files:
        for sheet_name in sheet_names:
            df = dfs[sheet_name]
            
            # Find all columns that contain 'date' in the column name
            date_cols = [col for col in df.columns if 'date' in col.lower()]
            
            for col in date_cols:
                for idx, value in df[col].iteritems():
                    try:
                        # Ensure the column is in datetime format
                        df.at[idx, col] = pd.to_datetime(value, dayfirst=True)
                        
                        # Format the datetime column to the desired format
                        df.at[idx, col] = df.at[idx, col].strftime('%Y-%m-%dT%H:%M:%SZ')
                    except Exception as e:
                        # Find the Excel-like cell address
                        cell_address = f"{pd.io.formats.excel.ExcelFormatter()._format_col_num(idx + 1)}{idx + 2}"  # +1 for Excel 1-based index, +2 for header row
                        error_message = (f"Error in sheet '{sheet_name}', cell '{cell_address}', "
                                         f"column '{col}', row {idx + 2}: "
                                         f"{type(value)} value '{value}' - {str(e)}")
                        logging.error(error_message)
                        st.error(f"Error in sheet '{sheet_name}', column '{col}', row {idx + 2}: {str(e)}")

            # Other conversions
            for column in ["Vintage PQ", "Vintage", "Year established PQ", "Fund ID PQ", "Company ID PQ", "Fund ID Sub Strategy PQ", "Quartile Rank PQ", "Preqin Quartile Rank PQ"]:
                if column in df.columns:
                    try:
                        df[column].fillna(-1, inplace=True)
                        df[column] = df[column].apply(lambda x: str(int(x)) if x != -1 else np.nan)
                    except Exception as e:
                        error_message = f"Error in sheet '{sheet_name}', column '{column}', data: {df[column].values} - {str(e)}"
                        logging.error(error_message)
                        st.error(f"Error in sheet '{sheet_name}', column '{column}': {str(e)}")

            file_name = str(date + " " + name + "_" + sheet_name + ".csv")
            csv_data = df.to_csv(index=False, encoding="utf-8-sig")
            zipped_files.writestr(file_name, csv_data.encode('utf-8-sig'))
    zip_io.seek(0)
    return zip_io

btn = st.button("Download CSVs")

if btn:
    try:
        dfs = combine(xlsxs, sheet_names)
        zip_io = zipsdd_csvs(dfs, name, date, sheet_names)
        tmp_download_link = st.download_button(
            label='Download CSVs',
            data=zip_io,
            file_name='dataframes.zip',
            mime='application/zip'
        )
    except Exception as e:
        logging.error(f"An error occurred: {str(e)}")
        st.error(f"An error occurred: {str(e)}")
