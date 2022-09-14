#############################################################
# Obtém os relatórios PDF do Amplimed, movendo-os para o
# Google Drive e atualizando a planilha de gerenciamento
# Lean Stay.
#############################################################

##################################
# BIBLIOTECAS
##################################
from seleniumwire import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from anticaptchaofficial.recaptchav2proxyless import*
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd
import requests
import json
import os.path
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import re
import sys
from os import chdir, getcwd, listdir
from bs4 import BeautifulSoup
from pandas import json_normalize
from googleapiclient.http import MediaFileUpload
from base64 import b64decode
from urllib.parse import urlencode
from dotenv import load_dotenv
load_dotenv()

##################################
# CONSTANTES E VARIÁVEIS GLOBAIS
##################################
ALWAYS_CONFIRM_BEFORE_PROCEED = os.getenv('LS_RELATORIOS_ALWAYS_CONFIRM_BEFORE_PROCEED')
ALWAYS_MANUALLY_SOLVE_CAPTCHA = os.getenv('LS_RELATORIOS_ALWAYS_MANUALLY_SOLVE_CAPTCHA')
ENVIRONMENT = os.getenv('LS_RELATORIOS_ENVIRONMENT')
SPREADSHEET_MANAGEMENT = {}
SPREADSHEET_MANAGEMENT['staging'] = os.getenv('LS_RELATORIOS_SPREADSHEET_MANAGEMENT_STAGING')
SPREADSHEET_MANAGEMENT['production'] = os.getenv('LS_RELATORIOS_SPREADSHEET_MANAGEMENT_PRODUCTION')
RANGE = os.getenv('LS_RELATORIOS_RANGE')
DRIVE_FOLDER = {}
DRIVE_FOLDER['staging'] = os.getenv('LS_RELATORIOS_DRIVE_FOLDER_STAGING')
DRIVE_FOLDER['production'] = os.getenv('LS_RELATORIOS_DRIVE_FOLDER_PRODUCTION')
AMPLIMED_LOGIN_URL = os.getenv('LS_RELATORIOS_AMPLIMED_LOGIN_URL')
AMPLIMED_LOGIN_EMAIL = os.getenv('LS_RELATORIOS_AMPLIMED_LOGIN_EMAIL')
AMPLIMED_LOGIN_PASSWORD = os.getenv('LS_RELATORIOS_AMPLIMED_LOGIN_PASSWORD')
ANTICAPTCHA_KEY = os.getenv('LS_RELATORIOS_ANTICAPTCHA_KEY')
ANTICAPTCHA_WEBSITE_KEY = os.getenv('LS_RELATORIOS_ANTICAPTCHA_WEBSITE_KEY')
AMPLIMED_AUTHORIZATION_KEY = None
AMPLIMED_HISTORICO_URL = os.getenv('LS_RELATORIOS_AMPLIMED_HISTORICO_URL')
AMPLIMED_RELATORIO_URL = os.getenv('LS_RELATORIOS_AMPLIMED_RELATORIO_URL')
AMPLIMED_USUCLIN = os.getenv('LS_RELATORIOS_AMPLIMED_USUCLIN')
WAIT_TIME_SECONDS = int(os.getenv('LS_RELATORIOS_WAIT_TIME_SECONDS'))
chromeBrowser = None
googleDriveService = None
googleSheetService = None

##################################
# FUNÇÕES AUXILIARES
##################################

# Analisa o PDF, checando se todos os campos críticos estão preenchidos
def analisar_prontuario_completo(link_completo):
    url = 'https://api.leansaude.com.br/v1/pdfContents.php?s=' + link_completo + '&origin=gdextended'
    
    response = requests.request("GET", url)
    
    pdf = response.json()
    pdf = json_normalize(pdf)
    
    campos_validacao = ['fields.Anamnese',
    'fields.Data da avaliação',
    'fields.Acomodação mais recente',
    'fields.CID',
    'fields.Há indicação de home care (internação domiciliar)?',
    'fields.Há pendências postergando a alta?',
    'fields.Acessou prontuário?',
    'fields.Acessou equipe médica?',
    'fields.Contato com familiares?']
    
    pdf = pdf[campos_validacao]
    pdf.dropna(inplace=True)
    completo = len(pdf)
    print(pdf)
    return completo

##################################
# Gera a base de prontuários tratada do Amplimed
def analisar_cadastro_anterior(idp):

    # @todo: vide issue https://github.com/leansaude/leansaude-rpa-relatorios/issues/1
    payload = 'codp='+ str(idp) + '&action=GET_HEADERS&ordem=DESC&codcon=0'
    headers = {
      'authorization': AMPLIMED_AUTHORIZATION_KEY,
      'Content-Type': 'application/x-www-form-urlencoded'
    }
    
    response = requests.request("POST", AMPLIMED_HISTORICO_URL, headers=headers, data=payload)
    #print(response)
    
    beautiful = response.text
    beautiful = beautiful.replace("\\n", "")
    beautiful = beautiful.replace("\n", "")
    beautiful = beautiful.replace("\\", "")
    
    soup2 = BeautifulSoup(beautiful,"html5lib")
    
    ides = []
    datas = []
    textos = []
    for tag in soup2.find_all('div',{'class': 'card-header','id':True}) :
        ide = tag['id']
        #print(ide)
        data = soup2.find('div', id = str(ide)).h4.text
        datas.append(data)
        texto = soup2.find('div', id = str(ide)).text
        #print(texto)
        textos.append(texto)
        #print(data)
        ide = re.sub('[^0-9]', '', ide)
        #print(ide)
        ides.append(ide)
    
    d = {'cod.prontuário': ides, 'Datas': datas,'Descritivo': textos}
    base_prontuarios = pd.DataFrame(data=d)
    med = {'Descritivo': textos}
    df_med = pd.DataFrame(data=med)
    df_med = df_med['Descritivo'].str.split(pat = 'Atendido por :',expand=True)
    df_med = df_med[1].str.split(pat = 'Enunciado:',expand=True)
    df_med = df_med[0].str.split(pat = '-',expand=True)
    df_med.rename(columns = {0:'Médico'}, inplace = True)
    df_med = df_med['Médico']
    base_prontuarios = base_prontuarios.merge(df_med,how='left', left_index=True, right_index=True)
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace(" de ", "/")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Janeiro", "01")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Fevereiro", "02")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Maru00e7o", "03")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Abril", "04")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Maio", "05")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Junho", "06")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Julho", "07")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Agosto", "08")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Setembro", "09")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Outubro", "10")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Novembro", "11")
    base_prontuarios['Datas'] = base_prontuarios['Datas'].str.replace("Dezembro", "12")
    base_prontuarios['Datas'] = pd.to_datetime(base_prontuarios['Datas'], format='%d/%m/%Y')
    base_prontuarios['Médico'] = base_prontuarios['Médico'].str.upper()
    base_prontuarios['primeiro_nome_medico'] = base_prontuarios['Médico'].str.split(None, 1,expand = True)[0]
    base_prontuarios['enunciado'] = base_prontuarios['Descritivo'].str.count('Enunciado')
        
    return base_prontuarios

##################################
# Move o PDF ao Google Drive
def subir_pdf_google_drive(idp, idpront, nome_completo, i):
    try:
        # obtém o conteúdo direto do PDF fazendo uma chamada à API Amplimed
        params = {}
        params['usuclin'] = AMPLIMED_USUCLIN
        params['codcon'] = idpront
        params_encoded = urlencode(params) + "&campos[]=infos&campos[]=anamnese&campos[]=habvida&campos[]=listpro&campos[]=conclusao&campos[]=plantrat&campos[]=docs&campos[]=prescr"

        pdfRequestResult = callAmplimedApi(AMPLIMED_RELATORIO_URL, 'POST', params_encoded)
        print('Obtido conteúdo do PDF')
        
        # decodifica e salva em disco o conteúdo do PDF
        pdfContentsEncoded = json.loads(pdfRequestResult)
        pdfContentsEncoded = pdfContentsEncoded['pdf'] #base64-encoded PDF contents
        pdfContentsBytes = b64decode(pdfContentsEncoded, validate=True)
        f = open("pdfs\\" + nome_completo + '_' + idpront + '.pdf', 'wb')
        f.write(pdfContentsBytes)
        f.close()
        print('Salvo: ' + "pdfs\\" + nome_completo + '_' + idpront + '.pdf')
        
        # Etapas para carregar o arquivo para a pasta correta do google Drive, liberando acesso e gerando o link Web
       
        # pasta do google drive para subir os arquivos
        folder_id = DRIVE_FOLDER[ENVIRONMENT]
        
        # chamada para subir arquivo na pasta do google Drive e gerar o id único do arquivo
        file_metadata = {'name': nome_completo + '_' + idpront + '.pdf','parents': [folder_id]}
        media = MediaFileUpload("pdfs\\" + nome_completo + '_' + idpront + '.pdf',
                                mimetype='application/pdf')
        # pylint: disable=maybe-no-member
        file = googleDriveService.files().create(body=file_metadata, media_body=media,
                                      fields='id').execute()
        
        file_id = file.get('id')
        
        print('Upload concluído do PDF no Google Drive')
        
        # liberar permissões de acesso do arquivo que subimos no Google Drive
        request_body = {
            'role': 'reader',
            'type': 'anyone'
        }
        
        response_permission = googleDriveService.permissions().create(
            fileId=file_id,
            body=request_body
        ).execute()
        
        print('Arquivo Liberado para Todos')
        
        # gerar o Link (URL) para acessar
        response_share_link = googleDriveService.files().get(
            fileId=file_id,
            fields='webViewLink'
        ).execute()
        link_completo = response_share_link.get('webViewLink')
        print(link_completo)
        link_completo = link_completo.replace("drivesdk", "sharing")
        print(link_completo)
        
        # validação dos campos preenchidos
        validacao = analisar_prontuario_completo(link_completo)
        print(validacao)
        if validacao > 0:

            print('Entrou no looping de preenchimento do Google Sheet')
            
            linha = int(i) + 2
            
            preencher_google_sheets(link_completo, "Visitas", "Q", linha)
            print('preencheu link')

            preencher_google_sheets(idpront, "Visitas", "AG", linha)
            print('preencheu número prontuário')

            preencher_google_sheets("Realizada", "Visitas", "K", linha)
            print('preencheu visita realizada')

        else:

            print('Falta campo a ser preeenhido no prontuário: ' + nome_completo + '_' + idpront )

            if ALWAYS_CONFIRM_BEFORE_PROCEED == 'SIM':
                userInput = input('Prosseguir? (s/n)')
                if userInput == 'n' :
                    sys.exit()
        
    except:
        print("Não foi possível obter o PDF do prontuário do(a) paciente " + str(nome_completo))

        if ALWAYS_CONFIRM_BEFORE_PROCEED == 'SIM':
            userInput = input('Prosseguir? (s/n)')
            if userInput == 'n' :
                sys.exit()

##################################
# Obtém a chave de autorização das APIs Amplimed e salva em AMPLIMED_AUTHORIZATION_KEY
def getAmplimedAuthorizationKey():
    global AMPLIMED_AUTHORIZATION_KEY
    global chromeBrowser

    openAmplimed()

    if AMPLIMED_AUTHORIZATION_KEY :
        return
    
    if not chromeBrowser :
        print('Erro: chromeBrowser não definido.')
        return

    # extrai o AMPLIMED_AUTHORIZATION_KEY
    for request in chromeBrowser.requests :
        if request.headers['authorization'] :
            AMPLIMED_AUTHORIZATION_KEY = request.headers['authorization']
            print('Obtido token para chamadas à API Amplimed')
            break

    print('AMPLIMED_AUTHORIZATION_KEY: ' + str(AMPLIMED_AUTHORIZATION_KEY))

##################################
# Abre Amplimed no Chrome e efetua login (se necessário)
def openAmplimed():
    global chromeBrowser

    # stop if Amplimed already open
    if chromeBrowser:
        return
    
    options = Options()
    #options.add_argument('--headless')
    options.add_argument('window-size=1500,900')
    chromeService = Service(ChromeDriverManager().install())
    chromeBrowser = webdriver.Chrome(options=options,service=chromeService)
    chromeBrowser.get(AMPLIMED_LOGIN_URL)
    time.sleep(10)

    if AMPLIMED_AUTHORIZATION_KEY:
        print('AMPLIMED_AUTHORIZATION_KEY já definida. Apenas abriu Chrome e navegou ao site do Amplimed, mas não irá efetuar login.')
        return
    
    print("Iniciando login no Amplimed")
    loginEmail = chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[1]/div/div/input')
    loginEmail.send_keys(AMPLIMED_LOGIN_EMAIL)
    
    loginPassword = chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[2]/div/div/input')
    loginPassword.send_keys(AMPLIMED_LOGIN_PASSWORD)

    # só executa anti-captcha se assim configurado
    if ALWAYS_MANUALLY_SOLVE_CAPTCHA != 'SIM' :
        print("Iniciando destravamento do Captcha")
        solver = recaptchaV2Proxyless()
        solver.set_verbose(1)
        solver.set_key(ANTICAPTCHA_KEY)
        solver.set_website_url(AMPLIMED_LOGIN_URL)
        solver.set_website_key(ANTICAPTCHA_WEBSITE_KEY)
        response = solver.solve_and_return_solution()

        if response != 0:
            print(response)
            chromeBrowser.execute_script(f"document.getElementById('g-recaptcha-response-100000').innerHTML = '{response}'")
            chromeBrowser.find_element(By.XPATH, '//*[@id="loginform"]/div[3]/div/button').click()
        else:
            print(solver.err_string)

        time.sleep(10)
    else : # ALWAYS_MANUALLY_SOLVE_CAPTCHA == 'SIM'
        print("--> AGUARDANDO 30 SEGUNDOS PARA EFETUAR LOGIN MANUAL NO AMPLIMED... <--")
        time.sleep(30)

    # navega para uma página que requeira alguma requisição POST contendo
    # o authorization header
    print('Navegando para agenda Amplimed para obter authorization header')
    wait = WebDriverWait(chromeBrowser, timeout=30)
    wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="navigation"]/ul/li[2]/a'))).click()
    time.sleep(10)

##################################
# Realiza uma chamada a um endpoint da API Amplimed
def callAmplimedApi(url, method, params):
    global AMPLIMED_AUTHORIZATION_KEY
    global chromeBrowser

    getAmplimedAuthorizationKey()
    
    if not AMPLIMED_AUTHORIZATION_KEY:
        sys.exit('Erro: AMPLIMED_AUTHORIZATION_KEY não definido.')

    if not chromeBrowser:
        sys.exit('Erro: chromeBrowser não definido.')
    
    request = '''var xhr = new XMLHttpRequest();
    xhr.open("''' + method + '''", "''' + url + '''", false);
    xhr.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');
    xhr.setRequestHeader('authorization', "''' + AMPLIMED_AUTHORIZATION_KEY + '''");
    xhr.send("''' + params + '''");
    return xhr.response;'''

    return chromeBrowser.execute_script(request)

##################################
#Inserir data de hoje do cadastro do cliente no Google Sheets
#   dado = informação a ser inderida
#   aba=str com nome da aba
#   coluna = str com letra da coluna ('D' por exemplo),
#   linha = int numero da linha a ser adicionada
def preencher_google_sheets(dado, aba, coluna, linha):
    
    linha_adicionar = str(aba) +'!' + str(coluna)+str(linha) 
    valores_adicionados = [[dado]]
    result = sheet.values().update(spreadsheetId=SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                                range=linha_adicionar,valueInputOption = "USER_ENTERED",
                                   body={"values":valores_adicionados}).execute()


##################################
# ROTINA PRINCIPAL
##################################

# Selecionar data de corte para as análises de prontuários
corte = pd.to_datetime('25/08/2022', format='%d/%m/%Y')
abrangencia = pd.Timedelta("2 days")

# Constrói googleDriveService e googleSheetService para reuso sempre que necessário
googleDriveService = build('drive', 'v3')
googleSheetService = build('sheets', 'v4')

# Obtém visitas com relatórios pendentes
sheet = googleSheetService.spreadsheets()
result = sheet.values().get(spreadsheetId=SPREADSHEET_MANAGEMENT[ENVIRONMENT],
                            range=RANGE).execute()
values = result.get('values', [])
print(values)
df = pd.DataFrame(values[1:], columns=values[0])
completa = df[['cod.prontuário']].drop_duplicates()
completa = completa.dropna()
completa = completa.loc[completa['cod.prontuário']!=""]
completa['realizado'] = 1
df = df.loc[(df['Status'] == 'Agendada') & (df['Link do Relatório Amplimed'] == '')]
df['Data da visita'] = pd.to_datetime(df['Data da visita'], format='%d/%m/%Y')
df['data_limite_tolerancia_inicial'] = pd.to_datetime(df['data_limite_tolerancia_inicial'], format='%d/%m/%Y')
df['Profissional'] = df['Profissional'].str.upper()
df['primeiro_nome_medico'] = df['Profissional'].str.split(None, 1, expand = True)[0]
df = df.loc[df['Data da visita']>=corte]

print('Total de relatórios pendentes: ' + str(len(df)))

if len(df) > 0:

    # para cada relatório pendente...
    for i in df.index:

        # obter parâmetros
        idp = df.loc[i,'ID Amplimed']
        nome_completo = df.loc[i,'Nome completo']
        nome_medico = df.loc[i,'primeiro_nome_medico']
        data_visita = df.loc[i,'Data da visita']
        abrangencia = df.loc[i,'data_limite_tolerancia_inicial']
        
        # realizar print no console para conferência
        print(nome_completo)
        print(idp)
        print(nome_medico)
        #print(data_visita)
        print('Linha Google Sheets: ' + str(i+2))
        
        # analisar quantidade de visitas pendentes do mesmo paciente
        qnt_visitas = len(df.loc[df['ID Amplimed']==idp])
        
        try:
            # gerar a base de prontuários do paciente
            base_prontuarios = analisar_cadastro_anterior(idp)
            base_prontuarios = base_prontuarios.loc[base_prontuarios['Datas']>= corte]
            
            # obter do paciente todos os prontuários que já foram criados em outras visitas
            base_prontuarios = base_prontuarios.merge(completa, how='left', on='cod.prontuário')
            base_prontuarios = base_prontuarios.loc[base_prontuarios['realizado']!=1]
            
            # obter todos os prontuários que não são do médico que está planejado a visita. Caso planeje a visita para 1 médico,
            # porém o prontuários esteja com nome de outro médico ele não identificará por causa disso.
            base_prontuarios = base_prontuarios.loc[base_prontuarios['primeiro_nome_medico']==nome_medico]
            
            # filtrar prontuários muito antigos com data abaixo da data previsivel de visita menos o parâmetro abrangência
            base_prontuarios = base_prontuarios.loc[base_prontuarios['Datas']>= data_visita - (data_visita - abrangencia)]
            
            # obter todos aqueles prontuários que não apresentam enunciado (falha na confecção do prontuário)
            base_prontuarios = base_prontuarios.loc[base_prontuarios['enunciado']==1]
            
            # reconfigurar os índices da base para que seja simples obter os dados novos
            base_prontuarios = base_prontuarios.reset_index(drop=True)
            
            # contagem da quantidade de prontuários que faltam para preencher
            qnt_prontuarios = len(base_prontuarios)
            #print(qnt_prontuarios)
            
            if qnt_prontuarios == 0:
                print('Não foi possível encontrar prontuários a serem adicionados')
                
            else:
                if qnt_visitas == 1:
                    idpront = base_prontuarios.iloc[0,:]['cod.prontuário']
                    print(idpront)
                   
                    subir_pdf_google_drive(idp, idpront, nome_completo, i)
                else:
                    print('Quantidade de visitas pendentes está maior que 1')
                
        except:
            print('Não há prontuários finalizados para o paciente ' + str(nome_completo))
        
        if ALWAYS_CONFIRM_BEFORE_PROCEED == 'SIM':
            userInput = input('Prosseguir? (s/n)')
            if userInput == 'n' :
                sys.exit()

        time.sleep(WAIT_TIME_SECONDS)