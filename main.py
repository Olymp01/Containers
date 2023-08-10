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

st.set_page_config(page_title='Контейнеры', page_icon=None, layout="centered", initial_sidebar_state="auto", menu_items=None)
st.header('Распечатка штрихкодов для контейнеров/пробирок')
log_name = st.text_input('Имя')
log_title = st.text_input('Логин')
log_pass = st.text_input('Пароль') 

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
rows = run_query(f'SELECT * FROM "{sheet_url}"')
df = pd.DataFrame(rows)
st.write(df)

def auth(sheet_url,new_gid):
    return sheet_url.replace("gid=0","gid="+new_gid)
    


