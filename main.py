import requests
import pandas as pd
import datetime
import shutil
import os
from dotenv import load_dotenv

# Obtendo a data atual
data = datetime.date.today()
ano = data.year
mes = data.month
dia = data.day

import os
import sys

# Verifica se estamos rodando como um executável PyInstaller
if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(application_path, 'template', 'modelo.xlsx')

# Carregando a planilha modelo
tabela = pd.read_excel(file_path, sheet_name="Plan3")

# Inicializando variáveis
iDATA = iTEXTMORNING = ""
iCHUVA = iUMIDADEMIN = iUMIDADEMAX = iTEMPERATURAMIN = iTEMPERATURAMAX = 0

# Token de acesso e ID da cidade (Sapucaia do Sul)
load_dotenv()
iTOKEN = os.getenv("iTOKEN")
iCIDADE = os.getenv("iCIDADE")

# Código do tipo da consulta
iTIPOCONSULTA = 3
try:
    # 1=Tempo agora na cidade
    if iTIPOCONSULTA == 1:
        iURL = f"http://apiadvisor.climatempo.com.br/api/v1/weather/locale/{iCIDADE}/current?token={iTOKEN}"
        iRESPONSE = requests.get(iURL)
        iRETORNO_REQ = iRESPONSE.json()

        for iCHAVE in iRETORNO_REQ:
            print(f"{iCHAVE} : {iRETORNO_REQ[iCHAVE]}")

        for iCHAVE in iRETORNO_REQ['data']:
            print(f"{iCHAVE} : {iRETORNO_REQ['data'][iCHAVE]}")

    # 2=Status do tempo no país
    if iTIPOCONSULTA == 2:
        iURL = f"http://apiadvisor.climatempo.com.br/api/v1/anl/synoptic/locale/BR?token={iTOKEN}"
        iRESPONSE = requests.get(iURL)
        iRETORNO_REQ = iRESPONSE.json()

        for iCHAVE in iRETORNO_REQ:
            print(f"País: {iCHAVE.get('country')}")
            print(f"Data: {iCHAVE.get('date')}")
            print(f"Descrição: {iCHAVE.get('text')}\n")

    # 3=Previsão para os próximos 15 dias
    if iTIPOCONSULTA == 3:
        iURL = f"http://apiadvisor.climatempo.com.br/api/v1/forecast/locale/{iCIDADE}/days/15?token={iTOKEN}"
        iRESPONSE = requests.get(iURL)
        iRETORNO_REQ = iRESPONSE.json()

        j = 0
        for iCHAVE in iRETORNO_REQ['data']:
            iDATA = iCHAVE.get('date_br')
            iUMIDADEMIN = iCHAVE['humidity']['min']
            iUMIDADEMAX = iCHAVE['humidity']['max']
            iTEXTMORNING = iCHAVE['text_icon']['text']['phrase']['reduced']
            iTEMPERATURAMIN = iCHAVE['temperature']['min']
            iTEMPERATURAMAX = iCHAVE['temperature']['max']

            tabela.loc[0+j, "temperatura"] = str(iTEMPERATURAMAX) + "ºC"
            tabela.loc[1+j, "temperatura"] = str(iTEMPERATURAMIN) + "ºC"
            tabela.loc[0+j, "umidade"] = str(iUMIDADEMAX) + "%"
            tabela.loc[1+j, "umidade"] = str(iUMIDADEMIN) + "%"

            # Adicionando as strings nas colunas apropriadas
            tabela.loc[0 + j, "descricao"] = iTEXTMORNING
            tabela.loc[0 + j, "data"] = iDATA

            j += 3

    # 6=Pesquisa ID da Cidade
    if iTIPOCONSULTA == 6:
        iCITY = input('Informe o nome da cidade: ')
        iURL = f"http://apiadvisor.climatempo.com.br/api/v1/locale/city?name={iCITY}&token={iTOKEN}"
        iRESPONSE = requests.get(iURL)
        iRETORNO_REQ = iRESPONSE.json()

        for iCHAVE in iRETORNO_REQ:
            iID = iCHAVE['id']
            iNAME = iCHAVE['name']
            iSTATE = iCHAVE['state']
            iCOUNTRY = iCHAVE['country']
            print(f"id: {iID} - state: {iSTATE} - country: {iCOUNTRY} - name: {iNAME}\n")

        iNEWCITY = input('Informe o ID da nova cidade ou 0 para sair: ')
        if iNEWCITY != "0":
            iURL = f"http://apiadvisor.climatempo.com.br/api-manager/user-token/{iTOKEN}/locales"
            payload = f"localeId[]={iNEWCITY}"
            headers = {'Content-Type': 'application/x-www-form-urlencoded'}
            iRESPONSE = requests.put(iURL, headers=headers, data=payload)
            print(iRESPONSE.text)
        else:
            exit()

    # Formatando a data para o nome do arquivo
    diaT = f"{dia:02d}"
    mesT = f"{mes:02d}"
    dataPlanilha = f"{diaT}.{mesT}.{ano}"


    # Criando a pasta de saída, se não existir
    os.makedirs("output", exist_ok=True)

    # Copiando o modelo de planilha para a pasta de saída
    shutil.copyfile(file_path, f"output/Previsão meteorológica {dataPlanilha}.xlsx")

    # Salvando os dados na planilha
    with pd.ExcelWriter(f"output/Previsão meteorológica {dataPlanilha}.xlsx", mode="a", if_sheet_exists='overlay') as escritor:
        tabela.to_excel(escritor, sheet_name="Plan3", index=False, merge_cells=False)

    print(f"Planilha 'Previsão meteorológica {dataPlanilha}.xlsx' criada com sucesso!")

except Exception as e:
  print(f"Erro: {e}")