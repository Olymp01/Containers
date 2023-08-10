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

options = {'Пробирка со средой Кэри Блера':'Z01','Пробирка со средой Эймса':'Z02', 'Уреазный тест(выдыхаемый воздух)':'Z03', 'Пробирка с желтой крышкой (ЦФДА)':'Z04',
          'Зеленая крышка без геля':'Z05','Конверт с ватной палочкой':'Z06','Баночка для кала':'Z07','Стерильный контейнер для мочи':'Z08',
           'Нестерильный контейнер для мочи':'Z09','Контейнер для мочи с иглой':'Z10','Контейнер для слюны':'Z11','Контейнер для спермы':'Z12',
           'Коричневая пробирка для мочи':'Z13','Мазок на степень чистоты':'Z14','Мазок отд. уретры':'Z15','РИФ':'Z16','Соскоб Эппендорф векторбест':'Z17',
           'Соскоб Эппендорф амплисенс':'Z18','Соскоб (цитология ПАП и шейки матки)':'Z19','Срез':'Z20','Стерильный контейнер для бакпосева':'Z21',
           'Тест полоска(тест антиген)':'Z22','Квантиферон 4 вакутейнера':'Z23','Перианальный соскоб(ректально)':'Z28','Пробирка "Streck" (trisomnia)':'Z29',
           'Соскоб ПЦР на коронавирус':'Z30','Микоплаза':'Z31','Фильтр-карта':'Z32'
          }

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
xd = pd.DataFrame(rows)
def auth(new_gid):
    return str(st.secrets["private_gsheets_url"]).replace("gid=0","gid="+str(new_gid))
    
for i in range(len(xd['Login'])):
    if log_title == xd['Login'][i] and log_pass == xd['Password'][i]:
        needed_sheet = auth(str(xd['Worksheet'][i]))
        rows2 = run_query(f'SELECT * FROM "{needed_sheet}"')
        q = pd.DataFrame(rows2)
        token = xd['Token'][i]
        st.write('Префикс: ',token)
        option = st.selectbox('Выбрать контейнер', options.keys())
        cont_prefix = options.get(option)   
        st.write(q[cont_prefix])     
        st.write(option,':',cont_prefix)
        number = st.number_input('Введите количество пробирок', min_value=0, step=1)
        st.write(number)
        
        coeff = 4
              
        def zapoln(cont_prefix):
            if cont_prefix in ('Z14','Z15','Z16','Z17','Z18'):
                return str('0000000')
            else:
                return str('0000000000')
            
        last = len(q[cont_prefix])-1
        while(pd.isnull(q[cont_prefix][last])==True or q[cont_prefix][last] == ''):
            last = last - 1
            if last < 0:
                last = last + 1 
                q[cont_prefix][last] = cont_prefix + zapoln(cont_prefix)
                coeff = 3
    
        last_pref = int('1'+((str(q[cont_prefix][last]))[3:]))
        lst = []      
        for i in range(1, number+1):
            temp = str(last_pref + i)[1:]
            lst.append(cont_prefix + temp)
        df = pd.DataFrame()
        df[option] = lst
        st.write(df)
