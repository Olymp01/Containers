import pandas as pd
import streamlit as st
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from gsheetsdb import connect

credentials = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
    ],
)
conn = connect(credentials=credentials)
        


