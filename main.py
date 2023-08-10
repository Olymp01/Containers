import pandas as pd
import streamlit as st
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from openpyxl import load_workbook
from datetime import datetime
import gsheetsdb
from gsheetsdb import connect

credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
conn = connect(credentials=credentials)

def run_query(query):
    rows = conn.execute(query, headers=1)
    rows = rows.fetchall()
    return rows

sheet_url = st.secrets["private_gsheets_url"]
workbook = load_workbook(sheet_url, data_only=True)
second_sheet = workbook.worksheets[1]
rows = run_query(f'SELECT * FROM "{sheet_url}"')
st.write(rows)
rows2 = run_query(f'SELECT * FROM "{second_sheet}"')
st.write(rows2)
df = pd.DataFrame(rows)
st.write(df)
