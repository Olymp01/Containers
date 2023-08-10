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
file_name = st.json(
    {
  "type": "service_account",
  "project_id": "containers-395304",
  "private_key_id": "3106f2d91613f9c3d9a18266a3ba3681d8c1d92b",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDL3DeTrmV0xn2H\nOkv6U/a8Mo4f0GKC+hThzooVnXx2XKc69x6LJUxffCcDSMS+JAIbhYJFHcbVSetQ\nqJau6fswwRGCykIyOE0phG4t651c9K8xqj2n8rsCOY8RaYKd+J1P7sLuKcAgiu0u\nrZhnC/qLG5/NmHFNBV3CNibeUpmrbS3EPqT2G/LSBJsqzMHwtS/zM0Kkt8lim7Aq\nxyhaGkLHrEhMOrFlNTXKEhLwM9q7tWHF3SNL/rEk3krSdXuZwkBHnT8bSPoT+eOZ\n4ixrsUpKaz6epZINTIXtRc5p87t8OlxebhX3wxkS2PL1iyxZ+8/fBb41WtYzfyzV\noW8ZmcHdAgMBAAECggEADRUtCeYOGRrokRDthvtJdoAbZ9uZ3tmev8kJ9TjTSnid\nZibYHR+VJvw9jV6SUGODhnGgemqz88SCfKgnwheTqRLS2rJmjM4ugYTABLlosr+k\ng0rvbwOgUGm84+YCCGedQjUf6bnghxYeAjoIxvWEAYER1ddmegYR+Ms8iQBYX/cM\nlP6jgyBlj7Z11o6qnvU52svwCdaShYG8iJVvqiyEAFxYAZg7mWpJgyVARxoVB+mZ\neAJUVCLqOMKD2mnDwth5kMFNBm8Ln/PbTvjyk///EX4fTao1Le7AQ9CTRCPrQ5BT\nASopCrFE9C9xNtQROvorDypg9uVP4rg5Qof0mS3pwQKBgQDk6XOEMGqD5t2ROjKt\nAk+HXz8fOqV9Y5Za44JhCNT3fYqCKk0QApnnhOv7hiwya/iAPy3fuXhSuxph+Fdf\nyrobRm88xCdQSG42e2uGFh+CwU1O4Sofz32T5UbiC8+YKxBpi4Fh/iGA0692GoSO\nuQmY9WK1Z/grBXy8r2/1/Vy/IQKBgQDj+9yYvzTBHSoMdvzW0XQgQ1YyQmtir1ba\nXeIGkp8ibzrrpBmQsDCMRzuShKhxIpYTJac9QRwFYh9BU+rJU3ncegp8YFF4rMSU\n+ZpBHT3rcCt4wak/ptcPLDtmll8s6jiISRwbRERGgQRqSNhAS5hm2y88TNDShMpP\nHOXWQAxXPQKBgEs1WDqqHZTQmSNJ17R3+eEkLfz4q76Spaee8Aspd68IlCVH+KG1\n+RLT3SR6ZLL8Pl0EZPIIYbvstTJTAYH7fUHQ1mNEnxYFdhB4ZE9dnWS8VcYAvHJP\nHZcl0TAbaU05eN70csHbWO6WniNTexHZQYn7tT6ctjCMUPs9OK+9WmeBAoGAeNKF\nuj27C40VN73niUE/tcl56PDiUE50TQ3sN2eFBo7EPxWcpt15HR6zJ5c+XZbiygru\ncrwKyZ+SyOBcUY33yyyyWfABvV5yYDFX2qJQqnGr2DdqJt2Yo+XhJSEUF42ZoEB+\nsMShGmxNlrY8RPbLMdd/VQmwsaDGRt8dv0n6QFECgYEAnryVXyLJSH/ioHwR4wN8\nvcW5JvBFkYZYpfsP+jsKvOgxYaek8rRV4jJQC5TAD1BdmoI3H/D/yb+DLzQav4Zk\n9bixwYz8vuvVJl1EYQQY/XGbptvwjxK1SYqCAt5bj/UTGYPKwhceHuQFbDlpMz/V\ngQmYAtUf1d87JiOXNUSBDis=\n-----END PRIVATE KEY-----\n",
  "client_email": "pythonsheet@containers-395304.iam.gserviceaccount.com",
  "client_id": "103228156097101825777",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/pythonsheet%40containers-395304.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

)
creds = ServiceAccountCredentials.from_json_keyfile_name(file_name,scope)
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
    
       
        





        


