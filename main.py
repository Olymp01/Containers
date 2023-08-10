import pandas as pd
import streamlit as st
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.oauth2 import service_account
from datetime import datetime
import gsheetsdb
from gsheetsdb import connect

scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
creds = ServiceAccountCredentials.from_service_account_info(
    (st.secrets["gcp_service_account"]), scope)
client = gspread.authorize(creds)
sheet = client.open('Containers').get_worksheet(0)
python_sheet = sheet.get_all_records()
headers = python_sheet.pop(0)
xd = pd.DataFrame(python_sheet, columns=headers)
st.write(xd)
