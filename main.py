import pandas as pd
import streamlit as st
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime



#Aвторизация
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
file_name = 'https://github.com/Olymp01/Containers/blob/main/client_key.json'
creds = ServiceAccountCredentials.from_json_keyfile_name(st.secrets["key"],scope)
client = gspread.authorize(creds)

sheet = client.open('Containers').get_worksheet(0)
python_sheet = sheet.get_all_records()
headers = python_sheet.pop(0)
xd = pd.DataFrame(python_sheet, columns=headers)

st.set_page_config(page_title = 'Контейнеры')
st.header('Распечатка штрихкодов для контейнеров/пробирок')
lst = []

log_name = st.text_input('Имя')
log_title = st.text_input('Логин')
log_pass = st.text_input('Пароль') 

#Список пробирок для распечатки штрихкодов
options = {'Пробирка со средой Кэри Блера':'Z01','Пробирка со средой Эймса':'Z02', 'Уреазный тест(выдыхаемый воздух)':'Z03', 'Пробирка с желтой крышкой (ЦФДА)':'Z04',
          'Зеленая крышка без геля':'Z05','Конверт с ватной палочкой':'Z06','Баночка для кала':'Z07','Стерильный контейнер для мочи':'Z08',
           'Нестерильный контейнер для мочи':'Z09','Контейнер для мочи с иглой':'Z10','Контейнер для слюны':'Z11','Контейнер для спермы':'Z12',
           'Коричневая пробирка для мочи':'Z13','Мазок на степень чистоты':'Z14','Мазок отд. уретры':'Z15','РИФ':'Z16','Соскоб Эппендорф векторбест':'Z17',
           'Соскоб Эппендорф амплисенс':'Z18','Соскоб (цитология ПАП и шейки матки)':'Z19','Срез':'Z20','Стерильный контейнер для бакпосева':'Z21',
           'Тест полоска(тест антиген)':'Z22','Квантиферон 4 вакутейнера':'Z23','Перианальный соскоб(ректально)':'Z28','Пробирка "Streck" (trisomnia)':'Z29',
           'Соскоб ПЦР на коронавирус':'Z30','Микоплаза':'Z31','Фильтр-карта':'Z32'
          }
#'Номер заявки':'',

# Действия после успешной авторизации
for i in range(len(xd['Login'])):
    if log_title == xd['Login'][i] and log_pass == xd['Password'][i]:
        worksheet_number = xd['Worksheet'][i]
        sheet2 = client.open('Containers').get_worksheet(worksheet_number)
        asd_sheet = sheet2.get_all_records()
        headers = asd_sheet.pop(0)
        q = pd.DataFrame(asd_sheet, columns=headers)
        st.write(q)
        token = xd['Token'][i]
        st.write('Префикс: ',token)
        option = st.selectbox('Выбрать контейнер', options.keys())
        cont_prefix = options.get(option)   
        st.write(option,':',cont_prefix)
        number = st.number_input('Введите количество пробирок', min_value=0, step=1)
        st.write(number)
        
        coeff = 4
        
        def zapoln(cont_prefix):
            if cont_prefix in ('Z14','Z15','Z16','Z17','Z18'):
                return str('0000000')
            else:
                return str('0000000000')
            
        last = len(q[option])-1
        while(pd.isnull(q[option][last])==True or q[option][last] == ''):
            last = last - 1
            if last < 0:
                last = last + 1 
                q[option][last] = cont_prefix + zapoln(cont_prefix)
                coeff = 3
    
        last_pref = int('1'+((str(q[option][last]))[3:]))
        for i in range(1, number+1):
            temp = str(last_pref + i)[1:]
            lst.append(cont_prefix + temp)
        df = pd.DataFrame()
        df[option] = lst
        st.write(df)
        
            
        def to_excel(df):
           output = BytesIO()
           writer = pd.ExcelWriter(output, engine='openpyxl')
           df.to_excel(writer, index=False, sheet_name='Sheet1') 
           writer.close()
           processed_data = output.getvalue()
           return processed_data
         
        def insert(option,q):
            
            cell_row = last + coeff
            cell_column = int(cont_prefix[1:])  
            cell = sheet2.cell(cell_row, cell_column)
            cell.value = lst[len(lst)-1]
            sheet2.update_cells([cell])
            sheet3 = client.open('Containers').get_worksheet(worksheet_number)
            asd_sheet = sheet3.get_all_records()
            headers = asd_sheet.pop(0)
            a = pd.DataFrame(asd_sheet, columns=headers)
            st.write(a)
            
        def run_cap():
            cap_button = st.button("Подтвердить") # Give button a variable name
            if cap_button: # Make button a condition.
                insert(option,q)
                st.text("Успешно внедрено")
        run_cap()   
        
        now = datetime.now()
        dt_string = now.strftime("%d.%m.%Y(%H-%M-%S)")
        df_xlsx = to_excel(df)
        st.download_button(label='📥 Скачать готовый файл',
                                       data = df_xlsx ,
                                       file_name= dt_string+'_'+option+'_'+log_title+'_'+log_name+'.xlsx') 
    
       
        





        

