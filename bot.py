"""
`pip install --upgrade -r requirements.txt`

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""
from botcity.core import DesktopBot
from google_credentials import Create_Service
from googleapiclient.http import MediaFileUpload
import pandas as pd
import os
import gspread

CLIENT_SECRET_FILE = 'credentials.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

bot = DesktopBot()
folder_id = '0Bxf69KA29_sEfk1keGxTSE4zeVpKMmVZUkFzcnVpNi1LM3FGMjZGN3VjU2ttOVFfekJYeG8'
folder_name = ''
query = f"parents = '{folder_id}' and name = '{folder_name}'"

def check_if_active_or_not(query):
    global folder_id
    response = service.files().list(q = query).execute()
    files = response.get('files')

    if files != []:
        df = pd.DataFrame(files)
        folder_id = df['id'][0]
    else:
        folder_id = '0Bxf69KA29_sELTBBZkoycjdlWnc'
        query = f"parents = '{folder_id}' and name = '{folder_name}'"
        find_drive_folder_id(query)

def find_drive_folder_id(query):
    global folder_id
    response = service.files().list(q = query).execute()
    files = response.get('files')

    if files != []:
        df = pd.DataFrame(files)
        folder_id = df['id'][0]
    else:
        create_folder()

def create_folder():
    global folder_id
    global folder_name

    file_metadata = {
        'name': folder_name,
        'parents': [folder_id],
        'mimeType': 'application/vnd.google-apps.folder'
    }
    created_folder = service.files().create(body=file_metadata).execute()
    folder_id = created_folder['id']

def save_pdf_mensal(empresa, mes, ano, fechamento):
    file_name = f'{empresa} - {ano} - {mes} - {fechamento}'
    tempo = 20000
    if fechamento == 'Razão':
        tempo = 50000

    bot.type_keys(['ctrl', 'd'])
    if fechamento == 'Razão':
        if bot.find('erro', matching=0.8, waiting_time=30000):
            bot.enter()
            bot.wait(1000)
    else:
        if bot.find('erro', matching=0.8, waiting_time=5000):
            bot.enter()
            bot.wait(1000)
    bot.type_key(file_name)
    if bot.find( "area_de_trabalho", matching=0.9, waiting_time=300000):
        bot.click()
    if bot.find( "fechamento", matching=0.9, waiting_time=10000):
        bot.double_click()
    if bot.find("salvar", matching=0.9, waiting_time=20000):
        bot.click()
    bot.wait(tempo)
    while bot.find( "gerando_pdfs", matching=0.9, waiting_time=10000):
        bot.wait(1000)
    bot.wait(2000)

def save_pdf_anual(empresa, ano, fechamento):
    file_name = f'{empresa} - {ano} - {fechamento}'

    bot.type_keys(['ctrl', 'd'])
    if bot.find('erro', matching=0.8, waiting_time=5000):
        bot.enter()
        bot.wait(1000)
    bot.type_key(file_name)
    if bot.find( "area_de_trabalho", matching=0.9, waiting_time=30000):
        bot.click()
    if bot.find( "fechamento", matching=0.9, waiting_time=10000):
        bot.double_click()
    if bot.find("salvar", matching=0.9, waiting_time=20000):
        bot.click()
    while bot.find( "gerando_pdfs", matching=0.9, waiting_time=10000):
        bot.wait(1000)
    bot.wait(2000)

def save_excel(empresa, mes, ano, fechamento):
    file_name = f'{empresa} - {ano} - {mes} - {fechamento}.xlsx'

    if bot.find("salvar_excel", matching=0.9, waiting_time=10000):
        bot.click()
    if bot.find('erro', matching=0.8, waiting_time=15000):
        bot.enter()
        bot.wait(1000)
    bot.type_key(file_name)
    if bot.find( "area_de_trabalho", matching=0.9, waiting_time=300000):
        bot.click()
    if not bot.find( "fechamento", matching=0.9, waiting_time=10000):
        not_found("fechamento")
    bot.double_click()
    if bot.find("salvar", matching=0.9, waiting_time=20000):
        bot.click()
    if bot.find("excel_open", matching=0.9, waiting_time=2000000):
        bot.alt_f4()
    bot.wait(2000)

def main():
    global folder_name
    global folder_id

    check_empresa = ''
    gc = gspread.service_account(filename='./service_account.json')
    sh = gc.open("Processo Fiscal e Contabil | Robô")
    worksheet = sh.get_worksheet(6)
    df = pd.DataFrame(worksheet.get_all_records())
    filtered_df = df.query("Processado != 's'")
    folder_path = "C:/Users/robo/Desktop/Fechamento"
    print(filtered_df)

    if not filtered_df.empty:
        docs = [f for f in os.listdir(folder_path)]

        for doc in docs:
            if os.path.isfile(f"{folder_path}/{doc}"):
                os.remove(f"{folder_path}/{doc}")
    if not filtered_df.empty:
        os.startfile('C:/Contabil/contabil.exe', arguments='/contabilidade')

        if bot.find("username", matching=0.97, waiting_time=30000):
            bot.kb_type('J4321')
            bot.enter()
        if bot.find( "login_await", matching=0.97, waiting_time=10000):
            bot.wait(40000)

        for index, row in filtered_df.iterrows():
            codigo_empresa = row['Cód Cliente']
            empresa = row['Nome (NÃO PREENCHER)']
            data_inicial = row['Data Inicial (dd/mm/aaaa)'].replace('/','')
            data_final = row['Data Final (dd/mm/aaaa)'].replace('/','')
            ano = data_final[-4:]
            mes = data_final[2:4]

            if check_empresa != empresa:
                bot.key_f8()
                if not bot.find("troca_empresa", matching=0.9, waiting_time=20000):
                    not_found("troca_empresa")
                if not bot.find("codigo", matching=0.9, waiting_time=20000):
                    not_found("codigo")
                bot.click()
                bot.kb_type(f'{codigo_empresa}')
                bot.enter()
                bot.wait(10000)

                if row['Consolidar'] == 'Sim':
                    bot.type_keys(['alt', 'c'])
                    bot.type_key('v')
                    bot.wait(2000)
                    bot.type_keys(['alt', 's'])
                    bot.type_keys(['alt', 'o'])

            if row['Relatório'] == 'Offboarding Anual':
                bot.type_keys(['alt', 'r'])
                bot.type_key('a')

                bot.tab()
                bot.type_key(data_inicial)
                bot.tab()
                bot.type_key(data_final)
                
                bot.type_keys(['alt', 'o'])
                if bot.find("ativo_diferente", matching=0.9, waiting_time=40000):
                    bot.enter()
                while bot.find( "processando", matching=0.9, waiting_time=5000):
                    bot.wait(1000)
                if bot.find("balanco", matching=0.9, waiting_time=10000):
                    save_pdf_anual(empresa, ano, 'Balanço')
                    bot.key_esc()
                elif bot.find("balanco-2", matching=0.9, waiting_time=10000):
                    save_pdf_anual(empresa, ano, 'Balanço')
                    bot.key_esc()
                bot.wait(2000)
                bot.key_esc()

                bot.type_keys(['alt', 'r'])
                bot.type_key('m')
                bot.type_key('d')

                bot.type_key(data_inicial)
                bot.tab()
                bot.type_key(data_final)

                bot.type_keys(['alt', 'o'])

                while bot.find( "processando", matching=0.9, waiting_time=5000):
                    bot.wait(1000)
                if bot.find("dre", matching=0.9, waiting_time=20000):
                    save_pdf_anual(empresa, ano, 'DRE')
                    bot.key_esc()
                elif bot.find("dre-2", matching=0.9, waiting_time=20000):
                    save_pdf_anual(empresa, ano, 'DRE')
                    bot.key_esc()
                bot.key_esc()

                folder_name = f'Syhus - {empresa}'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                check_if_active_or_not(query)
                folder_name = 'Contabilidade'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = '4. Offboarding'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = '2. Fechamento Anual'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = f'{ano}'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
            elif row['Relatório'] == 'Offboarding Mensal':
                bot.type_keys(['alt', 'r'])
                bot.type_key('r')

                bot.tab()
                bot.type_key(data_inicial)
                bot.tab()
                bot.type_key(data_final)

                bot.type_keys(['alt', 'o'])
                if bot.find("razao", matching=0.9, waiting_time=600000):
                    save_pdf_mensal(empresa, mes, ano, 'Razão')
                    save_excel(empresa, mes, ano, 'Razão')
                    bot.key_esc()
                bot.key_esc()

                bot.type_keys(['alt', 'r'])
                bot.type_key('b')
                bot.enter()

                bot.tab()
                bot.type_key(data_inicial)
                bot.tab()
                bot.type_key(data_final)

                if bot.find("balancete_mensal", matching=0.97, waiting_time=10000):
                    bot.click()

                bot.type_keys(['alt', 'o'])
                if bot.find("balancete", matching=0.9, waiting_time=30000):
                    save_pdf_mensal(empresa, mes, ano, 'Balancete Mensal')
                    save_excel(empresa, mes, ano, 'Balancete Mensal')
                    bot.key_esc()
                bot.wait(1000)
                bot.tab()
                bot.type_key(f'0101{ano}')
                bot.type_keys(['alt', 'o'])
                if bot.find("balancete", matching=0.9, waiting_time=30000):
                    save_pdf_mensal(empresa, mes, ano, 'Balancete Acumulado')
                    save_excel(empresa, mes, ano, 'Balancete Acumulado')
                    bot.key_esc()
                bot.wait(2000)
                bot.key_esc()

                bot.type_keys(['alt', 'r'])
                bot.type_key('m')
                bot.type_key('d')

                bot.type_key(data_inicial)
                bot.tab()
                bot.type_key(data_final)

                bot.type_keys(['alt', 'o'])
                if bot.find("dre", matching=0.9, waiting_time=30000):
                    save_pdf_mensal(empresa, mes, ano, 'DRE Mensal')
                    save_excel(empresa, mes, ano, 'DRE Mensal')
                    bot.key_esc()
                elif bot.find("dre-2", matching=0.9, waiting_time=10000):
                    save_pdf_mensal(empresa, mes, ano, 'DRE Mensal')
                    save_excel(empresa, mes, ano, 'DRE Mensal')
                    bot.key_esc()
                bot.wait(1000)
                bot.type_key(f'0101{ano}')
                bot.type_keys(['alt', 'o'])
                if bot.find("dre", matching=0.9, waiting_time=30000):
                    save_pdf_mensal(empresa, mes, ano, 'DRE Acumulado')
                    save_excel(empresa, mes, ano, 'DRE Acumulado')
                    bot.key_esc()
                elif bot.find("dre-2", matching=0.9, waiting_time=10000):
                    save_pdf_mensal(empresa, mes, ano, 'DRE Acumulado')
                    save_excel(empresa, mes, ano, 'DRE Acumulado')
                    bot.key_esc()
                bot.key_esc()

                folder_name = f'Syhus - {empresa}'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                check_if_active_or_not(query)
                folder_name = 'Contabilidade'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = '4. Offboarding'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = '1. Fechamento Mensal'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = f'{ano}'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)
                folder_name = f'{mes}-{ano}'
                query = f"parents = '{folder_id}' and name = '{folder_name}'"
                find_drive_folder_id(query)

            for file_name in os.listdir(folder_path):
                if file_name.endswith('.pdf'):
                    file_metadata = {
                        'name': file_name,
                        'parents': [folder_id]
                    }
                    service.files().create(
                        body=file_metadata,
                        media_body=MediaFileUpload(f"{folder_path}/{file_name}".format(file_name), mimetype="application/pdf"),
                        fields='id'
                    ).execute()
                elif file_name.endswith('.xlsx'):
                    file_metadata = {
                        'name': file_name,
                        'parents': [folder_id]
                    }
                    service.files().create(
                        body=file_metadata,
                        media_body=MediaFileUpload(f"{folder_path}/{file_name}".format(file_name), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                        fields='id'
                    ).execute()
                os.remove(f'{folder_path}/{file_name}')
            worksheet.update([['s']], f'G{index+2}')

            check_empresa = empresa

            folder_id = '0Bxf69KA29_sEfk1keGxTSE4zeVpKMmVZUkFzcnVpNi1LM3FGMjZGN3VjU2ttOVFfekJYeG8'
            folder_name = ''
        bot.alt_f4()
        bot.enter()

def not_found(label):
    print(f"Element not found: {label}")

if __name__ == '__main__':
    main()