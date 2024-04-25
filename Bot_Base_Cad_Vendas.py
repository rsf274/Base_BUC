#!/usr/bin/env python
# coding: utf-8

import time
import datetime as dt
import os
import re
import xlwings as xw
import pygetwindow as gw
import numpy as np
import pandas as pd
import glob
import pymsteams as teams
import functools
import operator
import tabulate

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

#===============================================CONFIGURAÇÃO DAS DATAS==================================================

ano_bissexto = (dt.datetime.today().year%4==0 and dt.datetime.today().year%100!=0) or (dt.datetime.today().year%400==0)
d = [31,29 if ((dt.datetime.today().month==2) and (ano_bissexto==True)) else 28,31,30,31,30,31,31,30,31,30,31][dt.datetime.today().month-1]

filtro_inicio_mes = dt.datetime.fromisoformat(f'{dt.datetime.today().year}-{dt.datetime.today().month:0>2}-{dt.datetime.today().replace(day=1).day:0>2}')
filtro_final_mes = dt.datetime.fromisoformat(f'{dt.datetime.today().year}-{dt.datetime.today().month:0>2}-{d:0>2}')

mes_ant = dt.datetime.today().replace(day=1,month=12,year=dt.datetime.today().year-1) if dt.datetime.today().month==1 else dt.datetime.today().replace(day=1,month=dt.datetime.today().month-1)
ano_bissexto_ant = (mes_ant.year%4==0 and mes_ant.year%100!=0) or (mes_ant.year%400==0)
d_ant = [31,29 if ((mes_ant.month==2) and (ano_bissexto==True)) else 28,31,30,31,30,31,31,30,31,30,31][mes_ant.month-1]
filtro_inicio_mes_ant = dt.datetime.fromisoformat(f'{mes_ant.year}-{mes_ant.month:0>2}-{mes_ant.replace(day=1).day:0>2}')
filtro_final_mes_ant = dt.datetime.fromisoformat(f'{mes_ant.year}-{mes_ant.month:0>2}-{d_ant:0>2}')

#=======================================================PASTAS==========================================================

pasta_arq_btx = 'C:\\Apuração de Resultados\\Pasta_btx'
pasta_financeiro = 'C:\\Financeiro'
pasta_csv_buc = 'C:\\Apuração de Resultados\\CSV_BUC'
pasta_base = 'C:\\Apuração de Resultados\\Base'
end_cnae = 'C:\\Apuração de Resultados\\CNAE'

#=============================================CONFIGURANDO O CANAL DO TEAMS===============================================

canal_teams = "{link webhook}"

msg_teams = teams.connectorcard(canal_teams)
msg_teams.text("# **Informativo do Laboratório de Dados**")
msg_teams.color('#00c0b6')

msg_teams1 = teams.connectorcard(canal_teams)
msg_teams1.text("# **Informativo do Laboratório de Dados**")
msg_teams1.color('#0B0965')

msg_teams2 = teams.connectorcard(canal_teams)
msg_teams2.text("# **Informativo do Laboratório de Dados** \n\n\n\n ## **►RANKINGS◄**")
msg_teams2.color('#0AF125')

#==========================================INICIAR NAVEGADOR E ACESSAR BTX==============================================

options = webdriver.ChromeOptions()
options.add_argument("--headless")
preferences = {"download.default_directory": "C:\Apuração de Resultados\Base\Arquivos Btx", "safebrowsing.enabled": "false"}
options.add_experimental_option("prefs", preferences)
driver = webdriver.Chrome(options = options, service = Service())

# Maximizando a tela
driver.maximize_window()

#Acessando o Site
driver.get("https://ic3.bitrix24.com.br/crm/deal/category/0/")
#Entrando com Username e OK
driver.find_element('xpath','/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[1]/div/input').send_keys('login')
time.sleep(2)
driver.find_element('xpath','/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[5]/button[1]').click()
time.sleep(5)

#Entrando com Senha e OK
driver.find_element('xpath','//*[@id="password"]').send_keys('senha')
time.sleep(2)
driver.find_element('xpath','/html/body/div[1]/div[2]/div/div[1]/div/div/div[3]/div/form/div/div[3]/button[1]').click()
time.sleep(30)    

# Todos os negócios
try:
    driver.get("https://ic3.bitrix24.com.br/crm/deal/list/")
    time.sleep(10)

except WebDriverException:
    driver.get("https://ic3.bitrix24.com.br/crm/deal/list/")
    time.sleep(10)

sel2 = driver.find_element('xpath','//*[@id="CRM_DEAL_LIST_V12_search"]')
sel2.click()

# Seleciona e aplicar filtro o Filtro BUC
buc = driver.find_element('xpath','//*[@id="popup-window-content-CRM_DEAL_LIST_V12_search_container"]/div/div/div[1]/div[2]/div[7]')
buc.click()
time.sleep(10)

# Clica na engrenagem de opções
driver.find_element('xpath','//*[@id="uiToolbarContainer"]/div[4]/button').click()
time.sleep(1)

# Seleciona "Exportar Negócios para CSV"
driver.find_element('xpath','//*[@id="popup-window-content-toolbar_deal_list_settings_menu"]/div/div/span[3]/span[2]').click()
time.sleep(2)

# Seleciona 'Exportar todos os campos do negócio' E 'Exportar SKU detalhadas'
driver.find_element('id','EXPORT_DEAL_CSV_opt_EXPORT_ALL_FIELDS_inp').click()
time.sleep(0.5)
driver.find_element('id','EXPORT_DEAL_CSV_opt_EXPORT_ALL_CLIENT_FIELDS_inp').click()
time.sleep(0.5)
driver.find_element('id','EXPORT_DEAL_CSV_opt_EXPORT_PRODUCT_FIELDS_inp').click()
time.sleep(0.5)
driver.find_element('css selector','#EXPORT_DEAL_CSV > div.popup-window-buttons > button.ui-btn.ui-btn-success.ui-btn-icon-start').click()
time.sleep(60)
driver.find_element('css selector','#EXPORT_DEAL_CSV > div.popup-window-buttons > a').click()
time.sleep(10)

os.chdir(pasta_arq_btx)
os.getcwd()

novo_nome = 'BTX.csv'

if os.path.exists(novo_nome):
    os.remove(novo_nome)
    
time.sleep(3)

list_of_files = glob.glob(f'{pasta_arq_btx}\\*')
arquivo = max(list_of_files , key=os.path.getctime)

os.replace(arquivo, novo_nome)
time.sleep(10)

driver.quit()

time.sleep(30)

#=================================================ARQUIVO BTX====================================================

BTX = pd.read_csv(f'{pasta_arq_btx}\\BTX.csv', sep=";", dtype=str, encoding='UTF-8')
lista_cad_pa = ['Empresa: Nome da Empresa','Empresa: Tipo de empresa',
                'Empresa: Telefone de trabalho','Empresa: Celular','Empresa: Email de trabalho',
                'Empresa: DOCUMENTO PA','Empresa: Endereço','Empresa: Complemento','Empresa: CEP',
                'Empresa: UF','Empresa: Bairro','Empresa: Cidade','Empresa: Agente de Expansão',
                'Empresa: CNAE PA','Empresa: Tipo de Pessoa','Empresa: Pessoa Responsável (CS)',
                'Empresa: Carteira','Data Processo Finalizado','Data de fechamento']
lista_deal = ['Pipeline','Fase','Renda','Empresa','Nome do negócio',
              'Criado','Modificado','Produto.1','Empresa: DOCUMENTO PA',
              'Empresa: Agente de Expansão','Último Status do Credenciamento']
lista_contatos = ['Contato: Primeiro Nome','Contato','Contato: Cargo','Contato: CPF',
                  'Empresa: Nome da Empresa','Empresa: Tipo de empresa','Empresa: DOCUMENTO PA']
lista_colunas = list(set(lista_cad_pa+lista_deal+lista_contatos))
BTX = BTX[lista_colunas].apply(lambda x: x.str.strip())
BTX = BTX[lista_colunas].apply(lambda x: x.str.upper())
BTX = BTX [['Empresa: DOCUMENTO PA','Empresa: Nome da Empresa','Empresa: Tipo de empresa',
          'Nome do negócio','Pipeline','Fase','Renda','Criado','Modificado','Produto.1',
          'Contato','Empresa: Telefone de trabalho','Empresa: Email de trabalho',
          'Empresa: Celular','Contato: Cargo','Contato: CPF','Empresa: Endereço',
          'Empresa: Complemento','Empresa: Bairro','Empresa: Cidade','Empresa: UF','Empresa: CEP',
          'Empresa: Agente de Expansão','Empresa: CNAE PA','Empresa: Tipo de Pessoa',
          'Empresa: Pessoa Responsável (CS)','Empresa: Carteira','Data Processo Finalizado',
          'Data de fechamento','Último Status do Credenciamento']]

#============================================ARRUMAR AGENTE DE EXPANSÃO===============================================

n_lin_btx = BTX.shape[0]
n_col_btx = BTX.shape[1]
colunas_btx = list(BTX.columns)
col_nome = colunas_btx.index('Empresa: Nome da Empresa')
col_agex = colunas_btx.index('Empresa: Agente de Expansão')

# REMOVE OS ACENTOS, CEDILHA E 'E' COMERCIAL DA COLUNA 'Empresa: Agente de Expansão'
BTX[colunas_btx[col_agex]] = BTX[colunas_btx[col_agex]].str.upper()
BTX['Empresa: Agente de Expansão'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
BTX['Empresa: Agente de Expansão'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
BTX['Empresa: Agente de Expansão'].replace(["Í","Ì"],"I", inplace = True, regex = True)
BTX['Empresa: Agente de Expansão'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
BTX['Empresa: Agente de Expansão'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
BTX['Empresa: Agente de Expansão'].replace("Ç","C", inplace = True, regex = True)

# REMOVE OS ACENTOS, CEDILHA, 'E' COMERCIAL E CPFS DA COLUNA 'Empresa: Nome da Empresa'
BTX['Empresa: Nome da Empresa'].replace([",","\.","-","\(","\)"]," ", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace(["Í","Ì"],"I", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
BTX['Empresa: Nome da Empresa'].replace("Ç","C", inplace = True, regex = True)
BTX[['Empresa: Nome da Empresa','Contato']] = BTX[['Empresa: Nome da Empresa','Contato']].apply(lambda x: x.str.rstrip("0123456789"))

# REMOVE OS AFIXOS DE PORTE/TIPO EMPRESARIAL DA COLUNA 'Empresa: Nome da Empresa'
BTX = BTX.to_numpy()
for i in range(n_lin_btx):
    if BTX[i][col_nome] != None:
        BTX[i][col_nome] = str(BTX[i][col_nome]).strip()
        BTX[i][col_nome] = re.compile("(\s(LTDA))?").sub("",str(BTX[i][col_nome]))
        BTX[i][col_nome] = re.compile("(\s((ME){2}))?").sub("",str(BTX[i][col_nome]))
        BTX[i][col_nome] = re.compile("(\s(S/S))?").sub("",str(BTX[i][col_nome]))
        BTX[i][col_nome] = re.compile("(\s(S/C))?").sub("",str(BTX[i][col_nome]))
        BTX[i][col_nome] = re.compile("(\s(EIRELI))?").sub("",str(BTX[i][col_nome]))
    if BTX[i][col_agex] != None:
        BTX[i][col_agex] = str(BTX[i][col_agex]).strip()
        BTX[i][col_agex] = re.compile("(\s(LTDA))?").sub("",str(BTX[i][col_agex]))
        BTX[i][col_agex] = re.compile("(\s((ME){2}))?").sub("",str(BTX[i][col_agex]))
        BTX[i][col_agex] = re.compile("(\s(S/S))?").sub("",str(BTX[i][col_agex]))
        BTX[i][col_agex] = re.compile("(\s(S/C))?").sub("",str(BTX[i][col_agex]))
        BTX[i][col_agex] = re.compile("(\s(EIRELI))?").sub("",str(BTX[i][col_agex]))
BTX = pd.DataFrame(BTX, columns = colunas_btx)
BTX = BTX.apply(lambda x: x.str.strip())

#==============================================ARRUMAR TIPO DE EMPRESA=================================================

BTX['CNPJ_PROC'] = BTX['Empresa: DOCUMENTO PA'].replace(["\.","\/","-"," "],"", regex = True)
BTX['Empresa: CNAE PA'] = BTX['Empresa: CNAE PA'].replace(["\.","-"," "],"", regex = True)
BTX['CNPJ_PROC'] = BTX['CNPJ_PROC'].str.strip()
BTX[['Criado','Modificado','Data Processo Finalizado','Data de fechamento']] = BTX[['Criado','Modificado','Data Processo Finalizado','Data de fechamento']].apply(lambda x: pd.to_datetime(x, errors='ignore', dayfirst = True))
BTX['Renda'].fillna("0", inplace = True)
BTX['Renda'] = BTX['Renda'].astype(np.float64)
BTX['Empresa: Email de trabalho'] = BTX['Empresa: Email de trabalho'].str.lower()

#================================================LISTA DE AGRs BTX===================================================

lista_agrs_btx = BTX[~BTX['CNPJ_PROC'].isna()][['CNPJ_PROC','Empresa: DOCUMENTO PA','Contato','Empresa: Nome da Empresa','Contato: CPF','Empresa: Tipo de empresa']]
lista_agrs_btx.rename(columns={'Empresa: DOCUMENTO PA':'DOCUMENTO PA','Contato':'AGR','Empresa: Nome da Empresa':'PA','Contato: CPF':'CPF','Empresa: Tipo de empresa':'Tipo de Ponto'}, inplace = True)
lista_agrs_btx['CNPJ_PROC2'] = lista_agrs_btx['CNPJ_PROC']
lista_agrs_btx['CNPJ_PROC2'].fillna("0", inplace = True)
lista_agrs_btx['CNPJ_PROC2'] = lista_agrs_btx['CNPJ_PROC2'].replace('nan',"0", regex = True)
lista_agrs_btx['CNPJ_PROC2'] = lista_agrs_btx['CNPJ_PROC2'].astype(np.int64)
lista_agrs_btx = lista_agrs_btx[((lista_agrs_btx['CNPJ_PROC2'] != 0) & ~(lista_agrs_btx['AGR'].isna()))]
lista_agrs_btx.drop('CNPJ_PROC2', axis=1, inplace=True)
lista_agrs_btx.drop_duplicates(subset = ['AGR','PA'], keep = 'first', inplace = True, ignore_index = True)

#===================================================QTD DE AGRs======================================================

Qtd_AGRs = BTX[((BTX['Contato: Cargo'].str.contains('PROPRIETÁRIO/AGR|AGENTE DE REGISTRO', na= False)) & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Contato: Cargo']]
Qtd_AGRs.rename(columns={'Contato: Cargo':'Quantidade de AGRs'}, inplace = True)

#================================================VALOR DOS PRODUTOS===================================================

Valor_produtos = BTX[~(BTX['CNPJ_PROC'].isna())].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).sum()[['CNPJ_PROC','Renda']]
Valor_produtos.rename(columns={'Renda':'Valor dos Produtos'}, inplace = True)

#======================================================FASES=========================================================

# DEFINIÇÃO DAS FASES 'PROCESSO FINALIZADO', 'EM PROCESSO DE CREDENCIAMENTO' E 'CANCELAMENTO' E CONTAGEM DESTES STATUS POR CNPJ
Fase = BTX[['CNPJ_PROC','Empresa: Tipo de empresa','Fase']].copy()
fase_cancelado = ['DESCREDENCIAMENTO','DESISTÊNCIA (MOTIVO NÃO INFORMADO)','DESCREDENCIAMENTO FORÇADO','QUIS RESCISÃO NA CARTEIRA','FOI PARA OUTRA AR','NEGÓCIO PERDIDO','CANCELAMENTO DE CADASTRO']
Fase['Fases para Contagem'] = np.where(Fase['Fase'].isin(fase_cancelado),"CANCELAMENTO",
                                       np.where(Fase['Fase']=="PROCESSO FINALIZADO","PROCESSO FINALIZADO",
                                                "EM PROCESSO DE CREDENCIAMENTO"))

Fase_PV = Fase[Fase['Empresa: Tipo de empresa']=='PV']
Fase_dif_PV = Fase[~(Fase['Empresa: Tipo de empresa']=='PV')]

Fase_final_PV = Fase_PV[((Fase_PV['Fases para Contagem']=="PROCESSO FINALIZADO") & ~(Fase_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_final_PV.rename(columns={'Empresa: Tipo de empresa':'Processo finalizado'}, inplace = True)
Fase_cred_PV = Fase_PV[((Fase_PV['Fases para Contagem']=="EM PROCESSO DE CREDENCIAMENTO") & ~(Fase_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_cred_PV.rename(columns={'Empresa: Tipo de empresa':'Em processo de credenciamento'}, inplace = True)
Fase_canc_PV = Fase_PV[((Fase_PV['Fases para Contagem']=="CANCELAMENTO") & ~(Fase_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_canc_PV.rename(columns={'Empresa: Tipo de empresa':'Cancelamento'}, inplace = True)

Fase_final = Fase_dif_PV[((Fase_dif_PV['Fases para Contagem']=="PROCESSO FINALIZADO") & ~(Fase_dif_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_final.rename(columns={'Empresa: Tipo de empresa':'Processo finalizado'}, inplace = True)
Fase_cred = Fase_dif_PV[((Fase_dif_PV['Fases para Contagem']=="EM PROCESSO DE CREDENCIAMENTO") & ~(Fase_dif_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_cred.rename(columns={'Empresa: Tipo de empresa':'Em processo de credenciamento'}, inplace = True)
Fase_canc = Fase_dif_PV[((Fase_dif_PV['Fases para Contagem']=="CANCELAMENTO") & ~(Fase_dif_PV['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Empresa: Tipo de empresa']]
Fase_canc.rename(columns={'Empresa: Tipo de empresa':'Cancelamento'}, inplace = True)

Fase_final = pd.concat([Fase_final_PV,Fase_final], axis = 0, ignore_index = True)
Fase_cred = pd.concat([Fase_cred_PV,Fase_cred], axis = 0, ignore_index = True)
Fase_canc = pd.concat([Fase_canc_PV,Fase_canc], axis = 0, ignore_index = True)

#======================================================PRODUTOS=========================================================

Curso_AGR_Adic = BTX[((BTX['Produto.1']=='CURSO (AGR ADICIONAL)') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Curso_AGR_Adic.rename(columns={'Produto.1':'Curso (AGR Adicional)'}, inplace = True)
Pack_Econ = BTX[((BTX['Produto.1']=='PACK ECONÔMICO') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Econ.rename(columns={'Produto.1':'Pack Econômico'}, inplace = True)
Migracao = BTX[((BTX['Produto.1']=='MIGRAÇÃO') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Migracao.rename(columns={'Produto.1':'Migração'}, inplace = True)
Curso = BTX[((BTX['Produto.1']=='CURSO') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Curso.rename(columns={'Produto.1':'Curso'}, inplace = True)
Pack_Basic = BTX[((BTX['Produto.1']=='PACK BASIC') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Basic.rename(columns={'Produto.1':'Pack Basic'}, inplace = True)
Curso_Migracao = BTX[((BTX['Produto.1']=='CURSO (MIGRAÇÃO)') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Curso_Migracao.rename(columns={'Produto.1':'Curso (Migração)'}, inplace = True)
Pack_Contador = BTX[((BTX['Produto.1']=='PACK CONTADOR') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Contador.rename(columns={'Produto.1':'Pack Contador'}, inplace = True)
Pack_Essencial = BTX[((BTX['Produto.1']=='PACK ESSENCIAL') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Essencial.rename(columns={'Produto.1':'Pack Essencial'}, inplace = True)
Pack_Master = BTX[((BTX['Produto.1']=='PACK MASTER') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Master.rename(columns={'Produto.1':'Pack Master'}, inplace = True)
Pack_Gold = BTX[((BTX['Produto.1']=='PACK GOLD') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Gold.rename(columns={'Produto.1':'Pack Gold'}, inplace = True)
Pack_Premium = BTX[((BTX['Produto.1']=='PACK PREMIUM') & ~(BTX['CNPJ_PROC'].isna()))].groupby(by='CNPJ_PROC', as_index = False, dropna = False, sort = False).count()[['CNPJ_PROC','Produto.1']]
Pack_Premium.rename(columns={'Produto.1':'Pack Premium'}, inplace = True)

#================================================CONSULTAS DA RUN - PARTE 1 (AE|PE|AGR)===================================================

# LÊ AS PLANS (ABAS) 'AE', 'PE' E 'AGR' DA PLANILHA 'APURAÇÃO DE EMISSÕES 4.0 - RUN'
plans_run = ['AE','PE','AGR']
for i in range(len(plans_run)):
    if plans_run[i]=='AE':
        RUN_AE = pd.read_excel(f'{pasta_financeiro}\\APURAÇAO DE EMISSOES 4.0 - Run.xlsm', sheet_name=f'{plans_run[i]}', usecols=['Agente de Expansão','Gerente de Expansão'], dtype=str)
        RUN_AE = RUN_AE[~RUN_AE['Agente de Expansão'].isna()]
        RUN_AE = RUN_AE.apply(lambda x: x.str.strip())
        RUN_AE = RUN_AE.apply(lambda x: x.str.upper())
    if plans_run[i]=='PE':
        RUN_PE = pd.read_excel(f'{pasta_financeiro}\\APURAÇAO DE EMISSOES 4.0 - Run.xlsm', sheet_name=f'{plans_run[i]}', usecols=['PE'], dtype=str)
        RUN_PE = RUN_PE[~RUN_PE['PE'].isna()]
        RUN_PE = RUN_PE.apply(lambda x: x.str.strip())
        RUN_PE = RUN_PE.apply(lambda x: x.str.upper())
    if plans_run[i]=='AGR':
        RUN_AGR = pd.read_excel(f'{pasta_financeiro}\\APURAÇAO DE EMISSOES 4.0 - Run.xlsm', sheet_name=f'{plans_run[i]}', dtype=str)
        RUN_AGR = RUN_AGR.apply(lambda x: x.str.strip())
        RUN_AGR = RUN_AGR.apply(lambda x: x.str.upper())

# REMOVE OS ACENTOS, CEDILHA, 'E' COMERCIAL E HÍFEN DAS COLUNAS 'Agente de Expansão' E 'Gerente de Expansão'
RUN_AE['Agente de Expansão'].replace("-"," ", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_AE['Agente de Expansão'].replace("Ç","C", inplace = True, regex = True)

RUN_AE['Gerente de Expansão'].replace("-"," ", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_AE['Gerente de Expansão'].replace("Ç","C", inplace = True, regex = True)
RUN_AE['Tipo de Parceiro'] = "AGENTE DE EXPANSÃO"

n_lin_ae = RUN_AE.shape[0]
n_col_ae = RUN_AE.shape[1]
colunas_run_ae = list(RUN_AE.columns)
col_ae = colunas_run_ae.index('Agente de Expansão')
col_ge = colunas_run_ae.index('Gerente de Expansão')

# REMOVE OS AFIXOS DE PORTE/TIPO EMPRESARIAL DAS COLUNAS 'Agente de Expansão' E 'Gerente de Expansão'
RUN_AE = RUN_AE.to_numpy()
for i in range(n_lin_ae):
    if RUN_AE[i][col_ae] != None:
        RUN_AE[i][col_ae] = str(RUN_AE[i][col_ae]).strip()
        RUN_AE[i][col_ae] = re.compile("(\s(LTDA))?").sub("",str(RUN_AE[i][col_ae]))
        RUN_AE[i][col_ae] = re.compile("(\s((ME){2}))?").sub("",str(RUN_AE[i][col_ae]))
        RUN_AE[i][col_ae] = re.compile("(\s(S/S))?").sub("",str(RUN_AE[i][col_ae]))
        RUN_AE[i][col_ae] = re.compile("(\s(S/C))?").sub("",str(RUN_AE[i][col_ae]))
        RUN_AE[i][col_ae] = re.compile("(\s(EIRELI))?").sub("",str(RUN_AE[i][col_ae]))
    if RUN_AE[i][col_ge] != None:
        RUN_AE[i][col_ge] = str(RUN_AE[i][col_ge]).strip()
        RUN_AE[i][col_ge] = re.compile("(\s(LTDA))?").sub("",str(RUN_AE[i][col_ge]))
        RUN_AE[i][col_ge] = re.compile("(\s((ME){2}))?").sub("",str(RUN_AE[i][col_ge]))
        RUN_AE[i][col_ge] = re.compile("(\s(S/S))?").sub("",str(RUN_AE[i][col_ge]))
        RUN_AE[i][col_ge] = re.compile("(\s(S/C))?").sub("",str(RUN_AE[i][col_ge]))
        RUN_AE[i][col_ge] = re.compile("(\s(EIRELI))?").sub("",str(RUN_AE[i][col_ge]))
RUN_AE = pd.DataFrame(RUN_AE, columns = colunas_run_ae)
RUN_AE = RUN_AE.apply(lambda x: x.str.strip())
RUN_AE.rename(columns={'Agente de Expansão':'Parceiro','Gerente de Expansão':'GE'}, inplace = True)

# REMOVE OS ACENTOS, CEDILHA, 'E' COMERCIAL, HÍFEN E CPFS DA COLUNA 'Ponto de Expansão'
RUN_PE['PE'].replace(["-","–"]," ", inplace = True, regex = True)
RUN_PE['PE'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_PE['PE'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_PE['PE'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_PE['PE'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_PE['PE'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_PE['PE'].replace("Ç","C", inplace = True, regex = True)
RUN_PE['PE'] = RUN_PE['PE'].str.rstrip("0123456789")
RUN_PE['GE'] = ""
RUN_PE['Tipo de Parceiro'] = "PONTO DE EXPANSÃO"

n_lin_pe = RUN_PE.shape[0]
n_col_pe = RUN_PE.shape[1]
colunas_run_pe = list(RUN_PE.columns)
col_pe = colunas_run_pe.index('PE')

# REMOVE OS AFIXOS DE PORTE/TIPO EMPRESARIAL DA COLUNA 'PE'
RUN_PE = RUN_PE.to_numpy()
for i in range(n_lin_pe):
    if RUN_PE[i][col_pe] != None:
        RUN_PE[i][col_pe] = str(RUN_PE[i][col_pe]).strip()
        RUN_PE[i][col_pe] = re.compile("(\s(LTDA))?").sub("",str(RUN_PE[i][col_pe]))
        RUN_PE[i][col_pe] = re.compile("(\s((ME){2}))?").sub("",str(RUN_PE[i][col_pe]))
        RUN_PE[i][col_pe] = re.compile("(\s(S/S))?").sub("",str(RUN_PE[i][col_pe]))
        RUN_PE[i][col_pe] = re.compile("(\s(S/C))?").sub("",str(RUN_PE[i][col_pe]))
        RUN_PE[i][col_pe] = re.compile("(\s(EIRELI))?").sub("",str(RUN_PE[i][col_pe]))
RUN_PE = pd.DataFrame(RUN_PE, columns = colunas_run_pe)
RUN_PE = RUN_PE.apply(lambda x: x.str.strip())

RUN_PE.drop_duplicates(subset = 'PE', keep = 'first', inplace = True, ignore_index = True)
RUN_PE['GE'] = np.where(RUN_PE['PE'] == "NE SOLUCOES","MATEUS",
                        np.where(RUN_PE['PE'] == "SD NEGOCIOS","SID", RUN_PE['GE']))
RUN_PE.rename(columns={'PE':'Parceiro'}, inplace = True)

# CONCATENA AS TABELAS PONTO DE EXPANSÃO E AGENTE DE EXPANSÃO, CRIANDO UMA TABELA ÚNICA E EXCLUINDO AS DUAS ANTERIORES
RUN_AE_PE = pd.concat([RUN_AE,RUN_PE], axis = 0, ignore_index = True)
del RUN_AE
del RUN_PE
RUN_AE_PE.drop_duplicates(subset = 'Parceiro', keep = 'first', inplace = True, ignore_index = True)
RUN_AE_PE['GE'].replace("nan","", inplace = True, regex = False)
RUN_AE_PE.to_csv(f'{pasta_csv_buc}\\AE_PE.csv', sep=",", decimal = ",", date_format = '%d/%m/%Y', index=False, encoding='UTF-8')

#=======================================================CNAES==========================================================

# LÊ A PLANILHA 'CNPJs_PAs', ONDE ESTÃO A RELAÇÃO DOS CLIENTES DA NOSSO CERTIFICADO E NOSSO SOLUÇÕES
cnae = pd.read_excel(f'{end_cnae}\\CNPJs_PAs.xlsx', sheet_name = "CNAE", usecols = ['CNPJ','Atividade CNAE'], dtype = str)
cnae = cnae.apply(lambda x: x.str.strip())
cnae['CNPJ'] = cnae['CNPJ'].replace('nan',"", regex = True)
cnae = cnae[~((cnae['CNPJ'].isna()) | (cnae['CNPJ']=='') | (cnae['Atividade CNAE'].isna()) | (cnae['Atividade CNAE']==''))]
cnae['CNPJ_PROC'] = cnae['CNPJ'].replace(["\.","\/","-"],"", regex = True)

#================================================CADASTRO PA===================================================

CADASTRO_PA = BTX[['CNPJ_PROC','Empresa: DOCUMENTO PA','Empresa: Nome da Empresa','Empresa: Tipo de empresa',
                      'Empresa: Telefone de trabalho','Empresa: Celular','Empresa: Email de trabalho','Criado',
                      'Empresa: Endereço','Empresa: Complemento','Empresa: CEP','Empresa: Bairro','Empresa: Cidade',
                      'Empresa: UF','Empresa: Agente de Expansão','Empresa: CNAE PA','Empresa: Tipo de Pessoa',
                      'Empresa: Pessoa Responsável (CS)','Empresa: Carteira','Data Processo Finalizado','Data de fechamento']].copy()
CADASTRO_PA = CADASTRO_PA[~((CADASTRO_PA['Empresa: Nome da Empresa'].str.contains('INATIV|Inati', na= False)) & (CADASTRO_PA['Empresa: Nome da Empresa'].isna()) & CADASTRO_PA['Empresa: DOCUMENTO PA'].isna())]
CADASTRO_PA.sort_values(by = ['Criado','CNPJ_PROC'], ascending = True, inplace = True, ignore_index = True)
CADASTRO_PA.drop_duplicates(subset = ['CNPJ_PROC','Empresa: Tipo de empresa'], keep = 'first', inplace = True, ignore_index = True)

CADASTRO_PA = CADASTRO_PA.join(RUN_AE_PE.set_index('Parceiro'), on = 'Empresa: Agente de Expansão', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(cnae.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')

CADASTRO_PA = CADASTRO_PA.join(Qtd_AGRs.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Valor_produtos.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Fase_cred.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Fase_final.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Fase_canc.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Curso_AGR_Adic.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Econ.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Migracao.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Curso.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Basic.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Curso_Migracao.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Contador.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Essencial.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Master.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Gold.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(Pack_Premium.set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix='_')

CADASTRO_PA[['Quantidade de AGRs','Em processo de credenciamento','Processo finalizado','Cancelamento','Curso (AGR Adicional)','Pack Econômico','Migração','Curso','Pack Basic','Curso (Migração)','Pack Contador','Pack Essencial','Pack Master','Pack Gold','Pack Premium']] = CADASTRO_PA[['Quantidade de AGRs','Em processo de credenciamento','Processo finalizado','Cancelamento','Curso (AGR Adicional)','Pack Econômico','Migração','Curso','Pack Basic','Curso (Migração)','Pack Contador','Pack Essencial','Pack Master','Pack Gold','Pack Premium']].apply(lambda x: x.replace(np.nan,0))
CADASTRO_PA[['Quantidade de AGRs','Em processo de credenciamento','Processo finalizado','Cancelamento','Curso (AGR Adicional)','Pack Econômico','Migração','Curso','Pack Basic','Curso (Migração)','Pack Contador','Pack Essencial','Pack Master','Pack Gold','Pack Premium']] = CADASTRO_PA[['Quantidade de AGRs','Em processo de credenciamento','Processo finalizado','Cancelamento','Curso (AGR Adicional)','Pack Econômico','Migração','Curso','Pack Basic','Curso (Migração)','Pack Contador','Pack Essencial','Pack Master','Pack Gold','Pack Premium']].apply(lambda x: x.astype(np.int64))

CADASTRO_PA['Valor dos Produtos'].fillna(0, inplace = True)
CADASTRO_PA['Valor dos Produtos'] = CADASTRO_PA['Valor dos Produtos'].astype(np.float64)

CADASTRO_PA['PA Ativo'] = np.where(CADASTRO_PA['Processo finalizado'] > 0,"SIM","NÃO")

CADASTRO_PA['PA em Credenciamento'] = np.where(((CADASTRO_PA['Em processo de credenciamento'] > 0) & (CADASTRO_PA['Processo finalizado'] == 0)),"SIM","NÃO")

CADASTRO_PA['PA Inativo'] = np.where(((CADASTRO_PA['Cancelamento'] == 0) & (CADASTRO_PA['Em processo de credenciamento'] == 0) & (CADASTRO_PA['Processo finalizado'] == 0)),"SIM",
                                     np.where(((CADASTRO_PA['Cancelamento'] > 0) & (CADASTRO_PA['Em processo de credenciamento'] == 0) & (CADASTRO_PA['Processo finalizado'] == 0)), "SIM", "NÃO"))

CADASTRO_PA['Total de Produtos'] = CADASTRO_PA[['Curso (AGR Adicional)','Pack Econômico','Migração','Curso','Pack Basic','Curso (Migração)','Pack Contador','Pack Essencial','Pack Master','Pack Gold','Pack Premium']].sum(axis = 1)

CADASTRO_PA['Data de fechamento'] = np.where(((CADASTRO_PA['PA Ativo'] == 'NÃO') & (CADASTRO_PA['Cancelamento'] > 0)), CADASTRO_PA['Data de fechamento'],pd.NaT)
CADASTRO_PA['Data de fechamento'] = pd.to_datetime(CADASTRO_PA['Data de fechamento'], errors='ignore', dayfirst = True)

CADASTRO_PA.rename(columns={'Empresa: DOCUMENTO PA':'DOCUMENTO PA','Empresa: Nome da Empresa':'Nome da Empresa','Empresa: Tipo de empresa':'Tipo da empresa','Empresa: Telefone de trabalho':'Telefone de trabalho',
                            'Empresa: Celular':'Celular','Empresa: Email de trabalho':'Email de trabalho','Empresa: Endereço':'Endereço','Empresa: Complemento':'Complemento','Empresa: Bairro':'Bairro',
                            'Empresa: Cidade':'Cidade','Empresa: UF':'UF','Empresa: CEP':'CEP','Empresa: Agente de Expansão':'Agente de Expansão','Empresa: CNAE PA':'CNAE PA',
                            'Empresa: Tipo de Pessoa':'Tipo de Pessoa','Empresa: Pessoa Responsável (CS)':'Pessoa Responsável (CS)','Empresa: Carteira':'Carteira','Data Processo Finalizado':'Data de Habilitação - PA',
                            'Data de fechamento':'Data de Inativação do PA','GE':'Gerente de Expansão'}, inplace = True)

CADASTRO_PA['CNPJ_PROC2'] = CADASTRO_PA['CNPJ_PROC']
CADASTRO_PA['CNPJ_PROC2'].fillna("0", inplace = True)
CADASTRO_PA['Agente de Expansão'] = CADASTRO_PA['Agente de Expansão'].replace('nan',"", regex = True)
CADASTRO_PA['CNPJ_PROC2'] = CADASTRO_PA['CNPJ_PROC2'].replace('nan',"0", regex = True)
CADASTRO_PA['CNPJ_PROC2'] = CADASTRO_PA['CNPJ_PROC2'].astype(np.int64)
CADASTRO_PA = CADASTRO_PA[(CADASTRO_PA['CNPJ_PROC2'] != 0)]
CADASTRO_PA.reset_index(inplace = True, drop = True)
CADASTRO_PA.drop('CNPJ_PROC2', axis=1, inplace=True)
CADASTRO_PA.sort_values(by='Criado', ascending = True, inplace = True, ignore_index = True)
CADASTRO_PA.drop_duplicates(subset = 'CNPJ_PROC', keep = 'first', inplace = True, ignore_index = True)

# MOSTRA A RELAÇÃO DE NOVOS PAS, COMPARANDO O ARQUIVO ANTERIOR COM A NOVA RELAÇÃO DE CADASTROS
base_cad_pa = pd.read_csv(f'{pasta_csv_buc}\\CADASTRO_PA.csv', sep=",", usecols = ['CNPJ_PROC','DOCUMENTO PA'], dtype = str, encoding='UTF-8')
base_cad_pa = base_cad_pa.apply(lambda x: x.str.strip())
novos_regs_pa = CADASTRO_PA[~CADASTRO_PA['CNPJ_PROC'].isin(base_cad_pa['CNPJ_PROC'])][['DOCUMENTO PA','Nome da Empresa','Criado','Tipo da empresa','Tipo de Pessoa','Carteira']]
novos_regs_pa.sort_values('Criado', ascending = True, inplace = True, ignore_index = True)
novos_regs_pa.rename(columns={'DOCUMENTO PA':'CNPJ'}, inplace = True)
novos_regs_pa.set_index('CNPJ', inplace = True)
if novos_regs_pa.shape[0]>0:
    novos_regs_pa.to_excel(f'{pasta_csv_buc}\\Aviso Teams\\Novos_PAs.xlsx', sheet_name = 'Novos PAs', index=False)
del base_cad_pa

#================================================CONSULTAS DA RUN - PARTE 2 (AGR)===================================================

# LIMPEZA E AJUSTES NAS COLUNAS 'AGR', 'PA' E 'COBRANÇA'
RUN_AGR['AGR'].replace("\*","", inplace=True, regex = True)
RUN_AGR['AGR'].replace("DIGTEC","", inplace=True, regex = True)
RUN_AGR['AGR'].replace("SISTEMA-","SISTEMA #", inplace=True, regex = True)

n_lin_agr = RUN_AGR.shape[0]
n_col_agr = RUN_AGR.shape[1]
colunas_run_agr = list(RUN_AGR.columns)
col_agr_agr = colunas_run_agr.index('AGR')
col_pa_agr = colunas_run_agr.index('PA')
col_cob_agr = colunas_run_agr.index('COBRANÇA')
col_docpa_agr = colunas_run_agr.index('DOCUMENTO PA')

RUN_AGR = RUN_AGR.to_numpy()
for i in range(n_lin_agr):
    if (re.compile("^(UNIDADES)\s?").search(str(RUN_AGR[i][col_pa_agr])) != None):
        RUN_AGR[i][col_pa_agr] = "VENDAS INTERNAS"
RUN_AGR = pd.DataFrame(RUN_AGR, columns = colunas_run_agr)

# REMOVE OS ACENTOS, CEDILHA, 'E' COMERCIAL, HÍFEN E CPFS DAS COLUNAS 'PA' E 'COBRANÇA'
RUN_AGR['PA'].replace(["-","\(","\)"]," ", inplace = True, regex = True)
RUN_AGR['PA'].replace([",","\.","\[","\]","  "],"", inplace = True, regex = True)
RUN_AGR['PA'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_AGR['PA'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_AGR['PA'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_AGR['PA'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_AGR['PA'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_AGR['PA'].replace("Ç","C", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["-","\(","\)"]," ", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace([",","\.","\[","\]","  "],"", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_AGR['COBRANÇA'].replace("Ç","C", inplace = True, regex = True)
RUN_AGR[['AGR','PA','COBRANÇA']] = RUN_AGR[['AGR','PA','COBRANÇA']].apply(lambda x: x.str.rstrip("0123456789"))

# REMOVE OS AFIXOS DE PORTE/TIPO EMPRESARIAL DAS COLUNAS 'AGR', 'PA', E 'COBRANÇA'
RUN_AGR = RUN_AGR.to_numpy()
for i in range(n_lin_agr):
    if RUN_AGR[i][col_agr_agr] != None:
        RUN_AGR[i][col_agr_agr] = str(RUN_AGR[i][col_agr_agr]).strip()
        RUN_AGR[i][col_agr_agr] = re.compile("(\s(LTDA))?").sub("",str(RUN_AGR[i][col_agr_agr]))
        RUN_AGR[i][col_agr_agr] = re.compile("(\s((ME){2}))?").sub("",str(RUN_AGR[i][col_agr_agr]))
        RUN_AGR[i][col_agr_agr] = re.compile("(\s(S/S))?").sub("",str(RUN_AGR[i][col_agr_agr]))
        RUN_AGR[i][col_agr_agr] = re.compile("(\s(S/C))?").sub("",str(RUN_AGR[i][col_agr_agr]))
        RUN_AGR[i][col_agr_agr] = re.compile("(\s(EIRELI))?").sub("",str(RUN_AGR[i][col_agr_agr]))
    if RUN_AGR[i][col_pa_agr] != None:
        RUN_AGR[i][col_pa_agr] = str(RUN_AGR[i][col_pa_agr]).strip()
        RUN_AGR[i][col_pa_agr] = re.compile("(\s(LTDA))?").sub("",str(RUN_AGR[i][col_pa_agr]))
        RUN_AGR[i][col_pa_agr] = re.compile("(\s((ME){2}))?").sub("",str(RUN_AGR[i][col_pa_agr]))
        RUN_AGR[i][col_pa_agr] = re.compile("(\s(S/S))?").sub("",str(RUN_AGR[i][col_pa_agr]))
        RUN_AGR[i][col_pa_agr] = re.compile("(\s(S/C))?").sub("",str(RUN_AGR[i][col_pa_agr]))
        RUN_AGR[i][col_pa_agr] = re.compile("(\s(EIRELI))?").sub("",str(RUN_AGR[i][col_pa_agr]))
    if RUN_AGR[i][col_cob_agr] != None:
        RUN_AGR[i][col_cob_agr] = str(RUN_AGR[i][col_cob_agr]).strip()
        RUN_AGR[i][col_cob_agr] = re.compile("(\s(LTDA))?").sub("",str(RUN_AGR[i][col_cob_agr]))
        RUN_AGR[i][col_cob_agr] = re.compile("(\s((ME){2}))?").sub("",str(RUN_AGR[i][col_cob_agr]))
        RUN_AGR[i][col_cob_agr] = re.compile("(\s(S/S))?").sub("",str(RUN_AGR[i][col_cob_agr]))
        RUN_AGR[i][col_cob_agr] = re.compile("(\s(S/C))?").sub("",str(RUN_AGR[i][col_cob_agr]))
        RUN_AGR[i][col_cob_agr] = re.compile("(\s(EIRELI))?").sub("",str(RUN_AGR[i][col_cob_agr]))
RUN_AGR = pd.DataFrame(RUN_AGR, columns = colunas_run_agr)

# REMOVE OS ACENTOS, CEDILHA, 'E' COMERCIAL, CARACTERES ESPECIAIS E CPFS DAS COLUNAS 'AGR' E 'PA'
RUN_AGR['AGR'].replace("INATIVO","", inplace = True, regex = True)
RUN_AGR['PA'].replace("INATIVO","", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["\.","-",","],"", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["Í","Ì"],"I", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
RUN_AGR['AGR'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
RUN_AGR['AGR'].replace("Ç","C", inplace = True, regex = True)
RUN_AGR['AGR'] = RUN_AGR['AGR'].str.rstrip("0123456789")

RUN_AGR['DOCUMENTO PA'] = RUN_AGR['DOCUMENTO PA'].astype(str)
RUN_AGR['DOCUMENTO PA'] = RUN_AGR['DOCUMENTO PA'].str.strip()
RUN_AGR['CNPJ_PROC'] = RUN_AGR['DOCUMENTO PA'].replace(["\.","\/","-"],"", regex = True)

RUN_AGR = RUN_AGR.join(CADASTRO_PA[['CNPJ_PROC','Tipo da empresa']].set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix = '_')
RUN_AGR['Tipo de Ponto'] = np.where(~((RUN_AGR['Tipo da empresa'].isna()) | (RUN_AGR['Tipo da empresa']=='')), RUN_AGR['Tipo da empresa'], RUN_AGR['Tipo de Ponto'])
RUN_AGR.drop('Tipo da empresa', axis=1, inplace=True)

RUN_AGR = RUN_AGR[~RUN_AGR['CNPJ_PROC'].isna()]
RUN_AGR = RUN_AGR.apply(lambda x: x.str.strip())
RUN_AGR.drop_duplicates(subset = ['AGR','PA'], keep = 'first', inplace = True, ignore_index = True)

lista_agrs_run = RUN_AGR[['CNPJ_PROC','DOCUMENTO PA','AGR','PA','COBRANÇA','CPF','Tipo de Ponto']].copy()

lista_agrs_run['CNPJ_PROC2'] = lista_agrs_run['CNPJ_PROC']
lista_agrs_run['CNPJ_PROC2'].fillna("0", inplace = True)
lista_agrs_run['CNPJ_PROC2'] = lista_agrs_run['CNPJ_PROC2'].replace('nan',"0", regex = True)
lista_agrs_run['CNPJ_PROC2'] = lista_agrs_run['CNPJ_PROC2'].astype(np.int64)

# lista_agrs_run['CNPJ_PROC'] = lista_agrs_run['CNPJ_PROC'].replace('nan',"0", regex = True)
# lista_agrs_run['CNPJ_PROC'] = lista_agrs_run['CNPJ_PROC'].astype(np.int64)
lista_agrs_run = lista_agrs_run[((lista_agrs_run['CNPJ_PROC2'] != 0) & ~(lista_agrs_run['AGR'].isna()))]
lista_agrs_run.drop('CNPJ_PROC2', axis=1, inplace=True)

lista_agrs_btx = lista_agrs_btx[~((lista_agrs_btx['AGR'].isin(lista_agrs_btx['AGR'])) & (lista_agrs_btx['PA'].isin(lista_agrs_btx['PA'])))]
lista_agrs_btx.reset_index(inplace = True, drop = True)
lista_agrs = pd.concat([lista_agrs_run,lista_agrs_btx], axis = 0, ignore_index = True)
lista_agrs.drop_duplicates(subset = ['AGR','PA'], keep = 'first', inplace = True, ignore_index = True)

lista_agrs2 = lista_agrs.copy()
lista_agrs2['AGR'].replace("SISTEMA #", "SISTEMA-", inplace=True, regex = True)
lista_agrs2['AGR'].replace("#", "", inplace=True, regex = True)

# MOSTRA A RELAÇÃO DE NOVOS AGRS, COMPARANDO O ARQUIVO ANTERIOR COM A NOVA RELAÇÃO DE CADASTROS
base_cad_agr = pd.read_csv(f'{pasta_csv_buc}\\AGRs.csv', sep=";", usecols = ['AGR','PA'], dtype = str, encoding='UTF-8')
base_cad_agr = base_cad_agr.apply(lambda x: x.str.strip())
novos_regs_agr = lista_agrs2[~lista_agrs2['AGR'].isin(base_cad_agr['AGR'])][['CPF','AGR','PA','COBRANÇA']]
novos_regs_agr['CPF'] = novos_regs_agr['CPF'].replace("nan","0", regex = True)
novos_regs_agr['CPF'].fillna("0", inplace = True)
novos_regs_agr['CPF2'] = novos_regs_agr['CPF'].replace(["\.","\/","-"],"", regex = True)
novos_regs_agr['CPF2'] = novos_regs_agr['CPF2'].astype(np.int64)
novos_regs_agr['CPF'] = np.where(novos_regs_agr['CPF2'] == 0,"---------",novos_regs_agr['CPF'])
novos_regs_agr.drop('CPF2', axis=1, inplace=True)
novos_regs_agr.sort_values('AGR', ascending = True, inplace = True, ignore_index = True)
novos_regs_agr.set_index('CPF', inplace = True)
if novos_regs_pa.shape[0]>0:
    novos_regs_agr.to_excel(f'{pasta_csv_buc}\\Aviso Teams\\Novos_AGRs.xlsx', sheet_name = 'Novos AGRs', index=False)
lista_agrs2.to_csv(f'{pasta_csv_buc}\\AGRs.csv', sep=";", index=False, encoding='UTF-8')
del lista_agrs2
del base_cad_agr

RUN_AGR_AGRs = RUN_AGR.drop_duplicates(subset = 'AGR', keep = 'last')[['AGR','CPF']]
RUN_AGR_PAs = RUN_AGR.drop_duplicates(subset = 'CNPJ_PROC', keep = 'last')[['CNPJ_PROC','PA','Tipo de Ponto']]
RUN_AGR_PAs['PA'] = np.where(RUN_AGR_PAs['CNPJ_PROC']=='00000000000100',"VENDAS INTERNAS",RUN_AGR_PAs['PA'])
RUN_AGR_PAs['Tipo de Ponto'] = np.where(RUN_AGR_PAs['CNPJ_PROC']=='00000000000100',"PA",RUN_AGR_PAs['Tipo de Ponto'])

filtro_pipeline = ['PÓS-VENDA AGR (NOVAS VENDAS)','PÓS-VENDA PV (NOVAS VENDAS)','PÓS-VENDA SHS (PV)','AGR-NOVO']
CONSULTA_NEGOCIOS_NOVO_BTX = BTX[BTX['Pipeline'].isin(filtro_pipeline)][['CNPJ_PROC','Pipeline','Fase','Empresa: DOCUMENTO PA','Empresa: Nome da Empresa','Contato','Criado','Renda','Produto.1','Modificado','Último Status do Credenciamento']].copy()
fase_cancelado = ['DESCREDENCIAMENTO','DESISTÊNCIA (MOTIVO NÃO INFORMADO)','DESCREDENCIAMENTO FORÇADO','QUIS RESCISÃO NA CARTEIRA','FOI PARA OUTRA AR','NEGÓCIO PERDIDO','CANCELAMENTO DE CADASTRO']
CONSULTA_NEGOCIOS_NOVO_BTX['Fases para Contagem'] = np.where(CONSULTA_NEGOCIOS_NOVO_BTX['Fase'].isin(fase_cancelado),"CANCELAMENTO",
                                                                np.where(CONSULTA_NEGOCIOS_NOVO_BTX['Fase']=="PROCESSO FINALIZADO","PROCESSO FINALIZADO",
                                                                         "EM PROCESSO DE CREDENCIAMENTO"))
CONSULTA_NEGOCIOS_NOVO_BTX = CONSULTA_NEGOCIOS_NOVO_BTX[~((CONSULTA_NEGOCIOS_NOVO_BTX['Empresa: Nome da Empresa'].isna())| (CONSULTA_NEGOCIOS_NOVO_BTX['Empresa: Nome da Empresa']=='') | (CONSULTA_NEGOCIOS_NOVO_BTX['Empresa: Nome da Empresa']=='nan'))]

CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'] = CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC']
CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'].fillna("0", inplace = True)
CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'] = CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'].replace('nan',"0", regex = True)
CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'] = CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'].astype(np.int64)
CONSULTA_NEGOCIOS_NOVO_BTX = CONSULTA_NEGOCIOS_NOVO_BTX[(CONSULTA_NEGOCIOS_NOVO_BTX['CNPJ_PROC2'] != 0)]
CONSULTA_NEGOCIOS_NOVO_BTX.drop('CNPJ_PROC2', axis=1, inplace=True)

CONSULTA_NEGOCIOS_NOVO_BTX = CONSULTA_NEGOCIOS_NOVO_BTX.join(CADASTRO_PA[['CNPJ_PROC','Agente de Expansão','Gerente de Expansão','Tipo de Parceiro']].set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CONSULTA_NEGOCIOS_NOVO_BTX['Criado'] = pd.to_datetime(CONSULTA_NEGOCIOS_NOVO_BTX['Criado'], errors='ignore', dayfirst = True)

CONSULTA_NEGOCIOS_NOVO_BTX.rename(columns={'Empresa: DOCUMENTO PA':'DOCUMENTO PA','Empresa: Nome da Empresa':'Nome da Empresa','Produto.1':'Produto'}, inplace = True)

CONSULTA_NEGOCIOS_NOVO_BTX.sort_values(by = ['CNPJ_PROC','Criado'], ascending = True, inplace = True, ignore_index = True)
n_lin_vendas = CONSULTA_NEGOCIOS_NOVO_BTX.shape[0]
for i in range(n_lin_vendas):
    if i==0:
        CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'Venda Nova'] = 1
    elif (CONSULTA_NEGOCIOS_NOVO_BTX.loc[i-1,'CNPJ_PROC']==CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'CNPJ_PROC']):
        CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'Venda Nova'] = 0
    elif (CONSULTA_NEGOCIOS_NOVO_BTX.loc[i-1,'CNPJ_PROC']!=CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'CNPJ_PROC']):
        CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'Venda Nova'] = 1
    else:
        CONSULTA_NEGOCIOS_NOVO_BTX.loc[i,'Venda Nova'] = 0
CONSULTA_NEGOCIOS_NOVO_BTX['Venda Adicional'] = np.where(CONSULTA_NEGOCIOS_NOVO_BTX['Venda Nova']==1,0,1)
CONSULTA_NEGOCIOS_NOVO_BTX[['Venda Nova','Venda Adicional']] = CONSULTA_NEGOCIOS_NOVO_BTX[['Venda Nova','Venda Adicional']].apply(lambda x: x.astype(np.int64))
CONSULTA_NEGOCIOS_NOVO_BTX.sort_values('Criado', ascending = False, inplace = True, ignore_index = True)

base_neg = pd.read_csv(f'{pasta_csv_buc}\\CONSULTA_NEGOCIOS_NOVO_BTX.csv', sep=";", usecols = ['DOCUMENTO PA','Criado','Venda Nova','Venda Adicional'], dtype = str, encoding='UTF-8')
base_neg = base_neg.apply(lambda x: x.str.strip())
base_neg[['Venda Nova','Venda Adicional']] = base_neg[['Venda Nova','Venda Adicional']].apply(lambda x: x.astype(np.int64))
base_neg['Criado'] = pd.to_datetime(base_neg['Criado'], errors='ignore', dayfirst = True)
base_neg = base_neg[((base_neg['Criado']>=filtro_inicio_mes) & (base_neg['Criado']<=filtro_final_mes))].reset_index(drop = True)
soma_vd_nova, soma_vd_add = base_neg['Venda Nova'].sum(),base_neg['Venda Adicional'].sum()
del base_neg

link_novos_pas = "[Novos PAs](https://nossoservicos-my.sharepoint.com/:x:/g/personal/apuracao3_nossoservicos_onmicrosoft_com/EeFvWCT-KW1Pt7u4kCci6wcBE6dtRrPEAW9mhtcRsBsjoA?e=1JnJxH)"
link_novos_agrs = "[Novos AGRs](https://nossoservicos-my.sharepoint.com/:x:/g/personal/apuracao3_nossoservicos_onmicrosoft_com/ERiA-tLerZpJk8E_-i2nrDIB-PgUgeCylsfxNF19VEGVnQ?e=oZTdVa)"
texto_nv_cads = []

card_teams = teams.cardsection()
card_teams.title(">CADASTROS PAs")
card_teams1 = teams.cardsection()
card_teams1.title(">CADASTROS AGRs")
card_teams2 = teams.cardsection()
card_teams2.title(">VENDAS")

if novos_regs_pa.shape[0]>0:
    card_teams.activityText(str(f"{novos_regs_pa.shape[0]} PAs novos cadastrados. \n\n")+link_novos_pas if novos_regs_pa.shape[0]>1 else str(f"{novos_regs_pa.shape[0]} PA novo cadastrado. \n\n")+link_novos_pas)
else:
    card_teams.activityText(str(f"""\nNenhum PA cadastrado. \n\n"""))
if novos_regs_agr.shape[0]>0:
    card_teams1.activityText(str(f"{novos_regs_agr.shape[0]} AGRs novos cadastrados. \n\n")+link_novos_agrs if novos_regs_agr.shape[0]>1 else str(f"{novos_regs_agr.shape[0]} AGR novo cadastrado. \n\n")+link_novos_agrs)
else:
    card_teams1.activityText(str(f"""\nNenhum AGR cadastrado. \n\n"""))
if ((soma_vd_nova>0) or (soma_vd_add>0)):
    card_teams2.activityText(str(f"## Vendas no mês {dt.datetime.today().month:0>2}/{dt.datetime.today().year}: \n\n\n\n")+str(f"- Novas: {soma_vd_nova} \r")+("\n\n")+str(f"- Adicionais: {soma_vd_add} \r"))
else:
    card_teams2.activityText(str(f"""\nNenhuma Venda no mês. \n\n"""))

msg_teams.addSection(card_teams)
msg_teams.addSection(card_teams1)
msg_teams.addSection(card_teams2)
msg_teams.send()

CONSULTA_NEGOCIOS_NOVO_BTX.to_csv(f'{pasta_csv_buc}\\CONSULTA_NEGOCIOS_NOVO_BTX.csv', sep=";", date_format = '%d/%m/%Y %H:%M:%S', index=False, encoding='UTF-8')

#================================================PACK CONTROLE DE VENDAS===================================================

PackControleDeVendas = pd.read_excel(f'{pasta_financeiro}\\Controle de Vendas C.xlsx', sheet_name='Pack',
                                     usecols=["UE","AE / PE","GE","PA","CNPJPA","AGR","PACK","PARCELA COMISSÃO",
                                              "VALOR TOTAL","DATA EMAIL","DATA FICHA","VENCIMENTO", "FORMA","COBRADO",
                                              "DESPESA","RECEBIDO","DATA REC","SITUAÇÃO ENVIOS","CÓD RASTREIO","CADASTRO",
                                              "CUSTO PACK","CUSTO NF","DIFERENÇA CUSTO","VALOR COMISSÃO AE","VALOR COMISSÃO GE",
                                              "VALOR COMISSÃO EX INTER","DATA PGMT COMISSÃO","Data 1º Pgto","Data 1º Venc"], dtype=str)
PackControleDeVendas = PackControleDeVendas.apply(lambda x: x.str.strip())
PackControleDeVendas.rename(columns={'CNPJPA':'CNPJ PA'}, inplace = True)
PackControleDeVendas = PackControleDeVendas[~(PackControleDeVendas['UE'].str.contains('CADASTRAR NA PLANILHA DE EMISSÕES', na=False))]
PackControleDeVendas = PackControleDeVendas[~((PackControleDeVendas['UE'].isna()) | (PackControleDeVendas['UE']==''))]

# Limpeza em PA
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.upper()
PackControleDeVendas['PA'].replace(["-",",","\.","\(","\)"]," ", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["Í","Ì"],"I", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
PackControleDeVendas['PA'].replace("Ç","C", inplace = True, regex = True)
PackControleDeVendas['PA'].replace(["INATIVO","  "], "", inplace=True, regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.rstrip("0123456789")
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.strip()

PackControleDeVendas['AGR'] = PackControleDeVendas['AGR'].str.upper()
PackControleDeVendas['AGR'].replace(["-",",","\.","\(","\)"]," ", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace(["Í","Ì"],"I", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
PackControleDeVendas['AGR'].replace("Ç","C", inplace = True, regex = True)
PackControleDeVendas['AGR'] = PackControleDeVendas['AGR'].str.strip()

PackControleDeVendas['AE / PE'] = PackControleDeVendas['AE / PE'].str.upper()
PackControleDeVendas['AE / PE'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
PackControleDeVendas['AE / PE'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
PackControleDeVendas['AE / PE'].replace(["Í","Ì"],"I", inplace = True, regex = True)
PackControleDeVendas['AE / PE'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
PackControleDeVendas['AE / PE'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
PackControleDeVendas['AE / PE'].replace("Ç","C", inplace = True, regex = True)
PackControleDeVendas['AE / PE'] = PackControleDeVendas['AE / PE'].str.strip()

PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("RFK ADMINISTRATIVO", "RFK DIGITAL E ADMINISTRATIVO", regex=False)
PackControleDeVendas['PA'] = PackControleDeVendas['PA'].str.replace("JOSEFO OLIVEIRA", "JOSEFA OLIVEIRA", regex=False)

n_lin_packs = PackControleDeVendas.shape[0]
n_col_packs = PackControleDeVendas.shape[1]
col_packs = list(PackControleDeVendas.columns)
col_ae_pe_packs = col_packs.index('AE / PE')
col_pa_packs = col_packs.index('PA')
col_cnpj_pa_packs = col_packs.index('CNPJ PA')
col_vt_packs = col_packs.index('VALOR TOTAL')
col_custo_packs = col_packs.index('CUSTO PACK')
col_custonf_packs = col_packs.index('CUSTO NF')
col_difcusto_packs = col_packs.index('DIFERENÇA CUSTO')
col_vcomae_packs = col_packs.index('VALOR COMISSÃO AE')
col_vcomge_packs = col_packs.index('VALOR COMISSÃO GE')
col_vcomexinter_packs = col_packs.index('VALOR COMISSÃO EX INTER')

PackControleDeVendas = PackControleDeVendas.to_numpy()
for i in range(n_lin_packs):
    if PackControleDeVendas[i][col_ae_pe_packs] != None:
        PackControleDeVendas[i][col_ae_pe_packs] = str(PackControleDeVendas[i][col_ae_pe_packs]).strip()
        PackControleDeVendas[i][col_ae_pe_packs] = re.compile("(\s(LTDA))?").sub("",str(PackControleDeVendas[i][col_ae_pe_packs]))
        PackControleDeVendas[i][col_ae_pe_packs] = re.compile("(\s((ME){2}))?").sub("",str(PackControleDeVendas[i][col_ae_pe_packs]))
        PackControleDeVendas[i][col_ae_pe_packs] = re.compile("(\s(S/S))?").sub("",str(PackControleDeVendas[i][col_ae_pe_packs]))
        PackControleDeVendas[i][col_ae_pe_packs] = re.compile("(\s(S/C))?").sub("",str(PackControleDeVendas[i][col_ae_pe_packs]))
        PackControleDeVendas[i][col_ae_pe_packs] = re.compile("(\s(EIRELI))?").sub("",str(PackControleDeVendas[i][col_ae_pe_packs]))
    if PackControleDeVendas[i][col_pa_packs] != None:
        PackControleDeVendas[i][col_pa_packs] = str(PackControleDeVendas[i][col_pa_packs]).strip()
        PackControleDeVendas[i][col_pa_packs] = re.compile("(\s(LTDA))?").sub("",str(PackControleDeVendas[i][col_pa_packs]))
        PackControleDeVendas[i][col_pa_packs] = re.compile("(\s((ME){2}))?").sub("",str(PackControleDeVendas[i][col_pa_packs]))
        PackControleDeVendas[i][col_pa_packs] = re.compile("(\s(S/S))?").sub("",str(PackControleDeVendas[i][col_pa_packs]))
        PackControleDeVendas[i][col_pa_packs] = re.compile("(\s(S/C))?").sub("",str(PackControleDeVendas[i][col_pa_packs]))
        PackControleDeVendas[i][col_pa_packs] = re.compile("(\s(EIRELI))?").sub("",str(PackControleDeVendas[i][col_pa_packs]))
    if (re.compile("NÃO TINHA UM AGR").search(str(PackControleDeVendas[i][col_pa_packs])) != None):
        PackControleDeVendas[i][col_cnpj_pa_packs] = "00.000.000/0000-00"
         
PackControleDeVendas = pd.DataFrame(PackControleDeVendas, columns = col_packs)

PackControleDeVendas2 = PackControleDeVendas[((PackControleDeVendas['CNPJ PA'].isna()) | (PackControleDeVendas['CNPJ PA'] == ""))]
PackControleDeVendas = PackControleDeVendas[~((PackControleDeVendas['CNPJ PA'].isna()) | (PackControleDeVendas['CNPJ PA'] == ""))]
PackControleDeVendas.reset_index(inplace = True, drop = True)
PackControleDeVendas2.reset_index(inplace = True, drop = True)

PackControleDeVendas2 = PackControleDeVendas2.join(CADASTRO_PA[['Nome da Empresa','DOCUMENTO PA']].set_index('Nome da Empresa'), on = 'PA', how = 'left', rsuffix = '_')
PackControleDeVendas1 = PackControleDeVendas2[~((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]
PackControleDeVendas2 = PackControleDeVendas2[((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]
PackControleDeVendas1.reset_index(inplace = True, drop = True)
PackControleDeVendas2.reset_index(inplace = True, drop = True)

if PackControleDeVendas2.shape[0]>0:
    for j in range(len(lista_agrs)):
        for i in range(len(PackControleDeVendas2)):
            valida = PackControleDeVendas2.loc[i,'PA']
            if (valida==lista_agrs.loc[j,'PA']):
                PackControleDeVendas2.loc[i,'DOCUMENTO PA'] = lista_agrs.loc[j,'DOCUMENTO PA']

PackControleDeVendas1 = pd.concat([PackControleDeVendas1,PackControleDeVendas2[~((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]], axis = 0, ignore_index = True)
PackControleDeVendas2 = PackControleDeVendas2[((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]
PackControleDeVendas2.reset_index(inplace = True, drop = True)

if PackControleDeVendas2.shape[0]>0:
    for j in range(len(lista_agrs)):
        for i in range(len(PackControleDeVendas2)):
            valida = PackControleDeVendas2.loc[i,'AGR']
            if (valida==lista_agrs.loc[j,'AGR']):
                PackControleDeVendas2.loc[i,'DOCUMENTO PA'] = lista_agrs.loc[j,'DOCUMENTO PA']

PackControleDeVendas1 = pd.concat([PackControleDeVendas1,PackControleDeVendas2[~((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]], axis = 0, ignore_index = True)
PackControleDeVendas2 = PackControleDeVendas2[((PackControleDeVendas2['DOCUMENTO PA'].isna()) | (PackControleDeVendas2['DOCUMENTO PA'] == ""))]
PackControleDeVendas2.reset_index(inplace = True, drop = True)

PackControleDeVendas1 = pd.concat([PackControleDeVendas1,PackControleDeVendas2], axis = 0, ignore_index = True)
PackControleDeVendas1['CNPJ PA'] = PackControleDeVendas1['DOCUMENTO PA']
PackControleDeVendas1['CNPJ PA'].fillna("00.000.000/0000-00", inplace = True)
PackControleDeVendas = pd.concat([PackControleDeVendas,PackControleDeVendas1], axis = 0, ignore_index = True)
del PackControleDeVendas1
del PackControleDeVendas2
PackControleDeVendas['CNPJ_PROC'] = PackControleDeVendas['CNPJ PA'].replace(["\.","\/","-"," "],"", regex = True)
PackControleDeVendas = PackControleDeVendas.join(CADASTRO_PA[['CNPJ_PROC','Criado']].set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix = '_')
PackControleDeVendas.drop(['DOCUMENTO PA'], axis=1, inplace=True)
PackControleDeVendas.sort_values('Criado', ascending = True, inplace = True, ignore_index = True)
PackControleDeVendas = PackControleDeVendas[['CNPJ_PROC','UE','AE / PE','GE','PA','CNPJ PA','AGR','PACK','PARCELA COMISSÃO','VALOR TOTAL','DATA EMAIL','DATA FICHA','VENCIMENTO','FORMA','COBRADO','DESPESA','RECEBIDO','DATA REC','SITUAÇÃO ENVIOS','CÓD RASTREIO','CADASTRO','CUSTO PACK','CUSTO NF','DIFERENÇA CUSTO','VALOR COMISSÃO AE','VALOR COMISSÃO GE','VALOR COMISSÃO EX INTER','DATA PGMT COMISSÃO','Data 1º Pgto','Data 1º Venc','Criado']]
PackControleDeVendas.to_csv(f'{pasta_csv_buc}\\PackControleDeVendas.csv', sep=";", decimal='.', index=False, encoding='UTF-8')

#================================================CURSOS AVULSOS===================================================

CursosAvulsosControleDeVendas = pd.read_excel(f'{pasta_financeiro}\\Controle de Vendas C.xlsx', sheet_name='Cursos Avulsos',
                                              usecols=['UE','AE / PE','RECEBIDO DE','CNPJ','AGR','PACK/CURSO','DATA EMAIL',
                                                       'DATA FICHA','FORMA PGMT','VENCIMENTO','PAGAMENTO','VALOR','RECEBIDO',
                                                       'AR','OBSERVAÇÕES','Valor em Aberto'], dtype=str)
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[~CursosAvulsosControleDeVendas['UE'].isna()]
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas.apply(lambda x: x.str.strip())
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas.apply(lambda x: x.str.upper())

# Limpeza em RECEBIDO DE
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["-",",","\.","\(","\)"],"", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.strip()

### Remover ####
n_lin_cursosavulsos = CursosAvulsosControleDeVendas.shape[0]
n_col_cursosavulsos = CursosAvulsosControleDeVendas.shape[1]
col_cursosavulsos = list(CursosAvulsosControleDeVendas.columns)
col_ae_pe_cursosavulsos = col_cursosavulsos.index('AE / PE')
col_rec_cursosavulsos = col_cursosavulsos.index('RECEBIDO DE')
col_agr_cursosavulsos = col_cursosavulsos.index('AGR')
col_cnpj_cursosavulsos = col_cursosavulsos.index('CNPJ')

CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas.to_numpy()
for i in range(n_lin_cursosavulsos):
    if CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] != None:
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]).strip()
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = re.compile("(\s(LTDA))?").sub("",str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = re.compile("(\s((ME){2}))?").sub("",str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = re.compile("(\s(S/S))?").sub("",str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = re.compile("(\s(S/C))?").sub("",str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos] = re.compile("(\s(EIRELI))?").sub("",str(CursosAvulsosControleDeVendas[i][col_ae_pe_cursosavulsos]))
    if CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] != None:
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]).strip()
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("(\s(LTDA))?").sub("",str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("(\s((ME){2}))?").sub("",str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("(\s(S/S))?").sub("",str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("(\s(S/C))?").sub("",str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("(\s(EIRELI))?").sub("",str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
    if CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] != None:
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("MATTEOS KELL").sub("MATEOS FERNANDO KELL", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("JOÃO COUTO").sub("JOÃO DO NASCIMENTO COUTO", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("MICHELLE").sub("MICHELLE STEPHANY", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("FERNANDA SILVA").sub("FERNANDA AMBROSINI SILVA", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("CICERO COSME ME").sub("CICERO COSME", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos] = re.compile("WEMERSON DELCONTI").sub("WEMERSON DEL COLI", str(CursosAvulsosControleDeVendas[i][col_agr_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] = re.compile("JC SILVA VENDAS").sub("J C DA SILVA VENDAS", str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos]))
    if (CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos] != None):
        CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos] = str(CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos]).strip()
        CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos] = str(np.where((re.compile("(\s+)?(PIEZO)(\s+)?").search(str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos])) != None),"00.000.000/0001-99",CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos] = str(np.where((re.compile("CF D AUDITORIA").search(str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos])) != None),"00.000.000/0001-98",CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos] = str(np.where((re.compile("J C DA SILVA VENDAS").search(str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos])) != None),"00.000.000/0001-97",CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos]))
        CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos] = str(np.where((re.compile("V CONTABILIDADE").search(str(CursosAvulsosControleDeVendas[i][col_rec_cursosavulsos])) != None),"00.000.000/0001-96",CursosAvulsosControleDeVendas[i][col_cnpj_cursosavulsos]))
CursosAvulsosControleDeVendas = pd.DataFrame(CursosAvulsosControleDeVendas, columns = col_cursosavulsos)

CursosAvulsosControleDeVendas['AE / PE'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AE / PE'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AE / PE'].replace(["Í","Ì"],"I", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AE / PE'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AE / PE'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AE / PE'].replace("Ç","C", inplace = True, regex = True)

CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["Í","Ì"],"I", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace("Ç","C", inplace = True, regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'].replace([".","  ","INATIVO"], "", inplace=True, regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.rstrip("0123456789")
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.strip()

CursosAvulsosControleDeVendas['AGR'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AGR'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AGR'].replace(["Í","Ì"],"I", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AGR'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AGR'].replace("Ç","C", inplace = True, regex = True)
CursosAvulsosControleDeVendas['AGR'].replace(["  ","INATIVO"], "", inplace=True, regex=False)
CursosAvulsosControleDeVendas['AGR'] = CursosAvulsosControleDeVendas['AGR'].str.strip()


# CORREÇÃO DE NA PLANILHA DE CURSOS
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("RFK ADMINISTRATIVO", "RFK DIGITAL E ADMINISTRATIVO", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("ANDERSON DE LIMASILVA", "ANDERSON LIMA SILVA", regex=False)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].str.replace("DR ASSESSORIA DOCUMENTAL", "D R ASSESSORIA DOCUMENTAL", regex=False)

CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[~(CursosAvulsosControleDeVendas['AGR'].isna())]

CursosAvulsosControleDeVendas2 = CursosAvulsosControleDeVendas[((CursosAvulsosControleDeVendas['CNPJ'].isna()) | (CursosAvulsosControleDeVendas['CNPJ'] == "") | (CursosAvulsosControleDeVendas['CNPJ']=='nan'))]
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[~((CursosAvulsosControleDeVendas['CNPJ'].isna()) | (CursosAvulsosControleDeVendas['CNPJ'] == "") | (CursosAvulsosControleDeVendas['CNPJ']=='nan'))]
CursosAvulsosControleDeVendas.reset_index(inplace = True, drop = True)
CursosAvulsosControleDeVendas2.reset_index(inplace = True, drop = True)

if CursosAvulsosControleDeVendas2.shape[0]>0:
    for j in range(len(lista_agrs)):
        for i in range(len(CursosAvulsosControleDeVendas2)):
            valida = str(CursosAvulsosControleDeVendas2.loc[i,'RECEBIDO DE'])
            if (valida==lista_agrs.loc[j,'PA']):
                CursosAvulsosControleDeVendas2.loc[i,'DOCUMENTO PA'] = lista_agrs.loc[j,'DOCUMENTO PA']

    CursosAvulsosControleDeVendas1 = CursosAvulsosControleDeVendas2[~((CursosAvulsosControleDeVendas2['DOCUMENTO PA'].isna()) | (CursosAvulsosControleDeVendas2['DOCUMENTO PA'] == "") | (CursosAvulsosControleDeVendas2['DOCUMENTO PA']=='nan'))]
    CursosAvulsosControleDeVendas2 = CursosAvulsosControleDeVendas2[((CursosAvulsosControleDeVendas2['DOCUMENTO PA'].isna()) | (CursosAvulsosControleDeVendas2['DOCUMENTO PA'] == "") | (CursosAvulsosControleDeVendas2['DOCUMENTO PA']=='nan'))]
    CursosAvulsosControleDeVendas2.reset_index(inplace = True, drop = True)

if CursosAvulsosControleDeVendas2.shape[0]>0:
    for j in range(len(lista_agrs)):
        for i in range(len(CursosAvulsosControleDeVendas2)):
            valida = CursosAvulsosControleDeVendas2.loc[i,'AGR']
            if (valida==lista_agrs.loc[j,'AGR']):
                CursosAvulsosControleDeVendas2.loc[i,'DOCUMENTO PA'] = lista_agrs.loc[j,'DOCUMENTO PA']

    CursosAvulsosControleDeVendas1 = pd.concat([CursosAvulsosControleDeVendas1,CursosAvulsosControleDeVendas2[~((CursosAvulsosControleDeVendas2['DOCUMENTO PA'].isna()) | (CursosAvulsosControleDeVendas2['DOCUMENTO PA'] == "") | (CursosAvulsosControleDeVendas2['DOCUMENTO PA']=='nan'))]], axis = 0, ignore_index = True)
    CursosAvulsosControleDeVendas2 = CursosAvulsosControleDeVendas2[((CursosAvulsosControleDeVendas2['DOCUMENTO PA'].isna()) | (CursosAvulsosControleDeVendas2['DOCUMENTO PA'] == "") | (CursosAvulsosControleDeVendas2['DOCUMENTO PA']=='nan'))]
    CursosAvulsosControleDeVendas2.reset_index(inplace = True, drop = True)

CursosAvulsosControleDeVendas1 = pd.concat([CursosAvulsosControleDeVendas1,CursosAvulsosControleDeVendas2], axis = 0, ignore_index = True)
CursosAvulsosControleDeVendas1['CNPJ'] = CursosAvulsosControleDeVendas1['DOCUMENTO PA']
CursosAvulsosControleDeVendas1['CNPJ'].fillna("00.000.000/0000-00", inplace = True)
CursosAvulsosControleDeVendas1['CNPJ'] = CursosAvulsosControleDeVendas1['CNPJ'].replace('nan',"00.000.000/0000-00", regex = True)
CursosAvulsosControleDeVendas1['CNPJ'] = CursosAvulsosControleDeVendas1['CNPJ'].astype(str)
CursosAvulsosControleDeVendas1['CNPJ'] = CursosAvulsosControleDeVendas1['CNPJ'].str.strip()
CursosAvulsosControleDeVendas = pd.concat([CursosAvulsosControleDeVendas,CursosAvulsosControleDeVendas1], axis = 0, ignore_index = True)
del CursosAvulsosControleDeVendas1
del CursosAvulsosControleDeVendas2
CursosAvulsosControleDeVendas['CNPJ_PROC'] = CursosAvulsosControleDeVendas['CNPJ'].replace(["\.","\/","-"," "],"", regex = True)
CursosAvulsosControleDeVendas['RECEBIDO DE'] = CursosAvulsosControleDeVendas['RECEBIDO DE'].replace('nan',"", regex = True)
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas.join(CADASTRO_PA[['CNPJ_PROC','Criado']].set_index('CNPJ_PROC'), on = 'CNPJ_PROC', how = 'left', rsuffix = '_')
CursosAvulsosControleDeVendas.drop(['DOCUMENTO PA'], axis=1, inplace=True)
CursosAvulsosControleDeVendas.sort_values('Criado', ascending = True, inplace = True, ignore_index = True)
CursosAvulsosControleDeVendas = CursosAvulsosControleDeVendas[['CNPJ_PROC','UE','AE / PE','RECEBIDO DE','CNPJ','AGR','PACK/CURSO','DATA EMAIL','DATA FICHA','FORMA PGMT','VENCIMENTO','PAGAMENTO','VALOR','RECEBIDO','AR','OBSERVAÇÕES','Valor em Aberto','Criado']]
CursosAvulsosControleDeVendas.to_csv(f'{pasta_csv_buc}\\CursosAvulsosControleDeVendas.csv', sep=";", index=False, encoding='UTF-8')

#================================================EMISSÕES===================================================

EMISSOES = pd.read_csv(f'{pasta_financeiro}\\APURAÇAO DE EMISSOES 4.0.csv', sep=",", dtype=str, encoding='UTF-8')
EMISSOES = EMISSOES.apply(lambda x: x.str.strip())
EMISSOES = EMISSOES.apply(lambda x: x.str.upper())
EMISSOES[['Vendedor','Cliente','A quem cobrar?']] = EMISSOES[['Vendedor','Cliente','A quem cobrar?']].apply(lambda x: x.str.upper())
# EMISSOES['Vendedor'].replace("\*","#", inplace=True, regex = True)
EMISSOES['Vendedor'].replace("\*","", inplace=True, regex = True)
EMISSOES['Vendedor'].replace("DIGTEC","", inplace=True, regex = True)
EMISSOES['Vendedor'].replace("SISTEMA-","SISTEMA #", inplace=True, regex = True)
EMISSOES['Vendedor'].replace("DIGTEC","", inplace=True, regex = True)
EMISSOES['A quem cobrar?'].replace("-CPF 888.888.88-88","", inplace=True, regex = True)
EMISSOES['A quem cobrar?'].replace("PATRICIA SILVA CPF 999.999.999-99","PATRICIA SILVA", inplace=True, regex = True)
EMISSOES['A quem cobrar?'].replace("LUCAS P. MARQUES","LUCAS P MARQUES", inplace=True, regex = True)
EMISSOES['AE ou PE'].replace(["J \& P PROMOTORA DE VENDAS LTDA","J\&P PROMOTORA DE VENDAS LTDA"],"J & P PROMOTORA DE VENDAS LTDA", inplace = True, regex = True)

EMISSOES['Vendedor'].replace(["-",",","\."],"", inplace = True, regex = True)
EMISSOES['Vendedor'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
EMISSOES['Vendedor'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
EMISSOES['Vendedor'].replace(["Í","Ì"],"I", inplace = True, regex = True)
EMISSOES['Vendedor'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
EMISSOES['Vendedor'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
EMISSOES['Vendedor'].replace("Ç","C", inplace = True, regex = True)
EMISSOES['Vendedor'].replace("INATIVO","", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["-","\(","\)"]," ", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace([",","\.","\[","\]","  "],"", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["Í","Ì"],"I", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
EMISSOES['A quem cobrar?'].replace("Ç","C", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace(["Á","À","Â","Ã"],"A", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace(["É","È","Ê","&"],"E", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace(["Í","Ì"],"I", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace(["Ó","Ò","Ô","Õ"],"O", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace(["Ú","Ù","Ü"],"U", inplace = True, regex = True)
EMISSOES['AE ou PE'].replace("Ç","C", inplace = True, regex = True)
EMISSOES[['Vendedor','A quem cobrar?','AE ou PE']] = EMISSOES[['Vendedor','A quem cobrar?','AE ou PE']].apply(lambda x: x.str.rstrip("0123456789"))

EMISSOES['Telefone'].replace([" ","\.","-","\(","\)"],"", inplace = True, regex = True)
EMISSOES[['Vendedor','A quem cobrar?','Telefone','AE ou PE']] = EMISSOES[['Vendedor','A quem cobrar?','Telefone','AE ou PE']].apply(lambda x: x.str.strip())

EMISSOES['Vendedor'].replace("^((FLAVIA FONTES)(#)?)", "FLAVIA FONTES#", inplace = True, regex = True)
EMISSOES['Vendedor'].replace("^((SCARLETT DE SOUZA)(#)?)", "SCARLETT DE SOUZA", inplace = True, regex = True)
EMISSOES['Vendedor'].replace("^((ANA JULIA MARTINELE)(#)?)", "ANA JULIA MARTINELE#", inplace = True, regex = True)
EMISSOES['Vendedor'].replace("^((KAMILA SOUZA E SILVA)(#)?)", "KAMILA SOUZA E SILVA", inplace = True, regex = True)

EMISSOES['PA_BUC'] = ""
EMISSOES['Tipo de empresa - Cadastro PA'] = ""
EMISSOES['CPF - AGR'] = ""
EMISSOES['Valida_cobrança'] = EMISSOES['A quem cobrar?']
EMISSOES['Valida_cobrança'] = EMISSOES['Valida_cobrança'].replace(r"(\sME)$","", regex = True)

n_lin_emissoes = EMISSOES.shape[0]
n_col_emissoes = EMISSOES.shape[1]
col_emissoes = list(EMISSOES.columns)
col_agr_emissoes = col_emissoes.index('Vendedor')
col_valida_emissoes = col_emissoes.index('Valida_cobrança')
col_ae_pe_emissoes = col_emissoes.index('AE ou PE')

EMISSOES = EMISSOES.to_numpy()
for i in range(n_lin_emissoes):
    if EMISSOES[i][col_agr_emissoes] != None:
        EMISSOES[i][col_agr_emissoes] = str(EMISSOES[i][col_agr_emissoes]).strip()
        EMISSOES[i][col_agr_emissoes] = re.compile("(\s(LTDA))?").sub("",str(EMISSOES[i][col_agr_emissoes]))
        EMISSOES[i][col_agr_emissoes] = re.compile("(\s((ME){2}))?").sub("",str(EMISSOES[i][col_agr_emissoes]))
        EMISSOES[i][col_agr_emissoes] = re.compile("(\s(S/S))?").sub("",str(EMISSOES[i][col_agr_emissoes]))
        EMISSOES[i][col_agr_emissoes] = re.compile("(\s(S/C))?").sub("",str(EMISSOES[i][col_agr_emissoes]))
        EMISSOES[i][col_agr_emissoes] = re.compile("(\s(EIRELI))?").sub("",str(EMISSOES[i][col_agr_emissoes]))
    if EMISSOES[i][col_valida_emissoes] != None:
        EMISSOES[i][col_valida_emissoes] = str(EMISSOES[i][col_valida_emissoes]).strip()
        EMISSOES[i][col_valida_emissoes] = re.compile("(\s(LTDA))?").sub("",str(EMISSOES[i][col_valida_emissoes]))
        EMISSOES[i][col_valida_emissoes] = re.compile("(\s((ME){2}))?").sub("",str(EMISSOES[i][col_valida_emissoes]))
        EMISSOES[i][col_valida_emissoes] = re.compile("(\s(S/S))?").sub("",str(EMISSOES[i][col_valida_emissoes]))
        EMISSOES[i][col_valida_emissoes] = re.compile("(\s(S/C))?").sub("",str(EMISSOES[i][col_valida_emissoes]))
        EMISSOES[i][col_valida_emissoes] = re.compile("(\s(EIRELI))?").sub("",str(EMISSOES[i][col_valida_emissoes]))
    if EMISSOES[i][col_ae_pe_emissoes] != None:
        EMISSOES[i][col_ae_pe_emissoes] = str(EMISSOES[i][col_ae_pe_emissoes]).strip()
        EMISSOES[i][col_ae_pe_emissoes] = re.compile("(\s(LTDA))?").sub("",str(EMISSOES[i][col_ae_pe_emissoes]))
        EMISSOES[i][col_ae_pe_emissoes] = re.compile("(\s((ME){2}))?").sub("",str(EMISSOES[i][col_ae_pe_emissoes]))
        EMISSOES[i][col_ae_pe_emissoes] = re.compile("(\s(S/S))?").sub("",str(EMISSOES[i][col_ae_pe_emissoes]))
        EMISSOES[i][col_ae_pe_emissoes] = re.compile("(\s(S/C))?").sub("",str(EMISSOES[i][col_ae_pe_emissoes]))
        EMISSOES[i][col_ae_pe_emissoes] = re.compile("(\s(EIRELI))?").sub("",str(EMISSOES[i][col_ae_pe_emissoes]))
EMISSOES = pd.DataFrame(EMISSOES, columns = col_emissoes)
EMISSOES = EMISSOES.apply(lambda x: x.str.strip())

EMISSOES['CNPJ PA_PROC'] = EMISSOES['CNPJ PA'].replace(["\.","\/","-"],"", regex = True)
EMISSOES['CNPJ PA_PROC'] = EMISSOES['CNPJ PA_PROC'].str.strip()

top_3_ae = EMISSOES[['AE ou PE','Conc IDAR']].groupby(by='AE ou PE',as_index=False,dropna=False).count()[['AE ou PE','Conc IDAR']]
top_3_ae.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
top_3_ae = top_3_ae[top_3_ae['AE ou PE'].isin(top_3_ae.loc[:2,'AE ou PE'])][['AE ou PE']]
top_3_ae = [top_3_ae.loc[i,'AE ou PE'] for i in range(len(top_3_ae))]

EMISSOES1 = EMISSOES[((EMISSOES['PA_BUC'].isna()) | (EMISSOES['PA_BUC']=='') | (EMISSOES['PA_BUC']=='nan'))]
EMISSOES = EMISSOES[~((EMISSOES['PA_BUC'].isna()) | (EMISSOES['PA_BUC']=='') | (EMISSOES['PA_BUC']=='nan'))]

EMISSOES_fora_top3_ae = EMISSOES1[~(EMISSOES1['AE ou PE'].isin(top_3_ae))]
EMISSOES_top3_ae = EMISSOES1[(EMISSOES1['AE ou PE'].isin(top_3_ae))]

EMISSOES_fora_top3_ae.reset_index(inplace = True, drop = True)
EMISSOES_top3_ae.reset_index(inplace = True, drop = True)

EMISSOES_fora_top3_ae['PA_BUC'] = EMISSOES_fora_top3_ae.join(RUN_AGR_PAs.set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['PA']]
EMISSOES_fora_top3_ae['Tipo de empresa - Cadastro PA'] = EMISSOES_fora_top3_ae.join(RUN_AGR_PAs.set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['Tipo de Ponto']]
EMISSOES_fora_top3_ae['CPF - AGR'] = EMISSOES_fora_top3_ae.join(RUN_AGR_AGRs.set_index('AGR'), on = 'Vendedor', how = 'left',rsuffix='_')[['CPF']]

EMISSOES_top3_ae['PA_BUC'] = EMISSOES_top3_ae.join(RUN_AGR_PAs.set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['PA']]
EMISSOES_top3_ae['Tipo de empresa - Cadastro PA'] = EMISSOES_top3_ae.join(RUN_AGR_PAs.set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['Tipo de Ponto']]
EMISSOES_top3_ae['CPF - AGR'] = EMISSOES_top3_ae.join(RUN_AGR_AGRs.set_index('AGR'), on = 'Vendedor', how = 'left',rsuffix='_')[['CPF']]

EMISSOES = pd.concat([EMISSOES,EMISSOES_fora_top3_ae,EMISSOES_top3_ae], axis = 0, ignore_index = True)

EMISSOES = EMISSOES.apply(lambda x: x.str.strip())
EMISSOES['E-mail'] = EMISSOES['E-mail'].str.lower()
EMISSOES['Tipo de empresa - Cadastro PA'].replace("", "INTERNO", inplace = True, regex = True)
EMISSOES['Tipo de empresa - Cadastro PA'].fillna("INTERNO", inplace = True)
EMISSOES['CPF - AGR'].replace("", "000.000.000-00", inplace = True, regex = True)
EMISSOES['CPF - AGR'].fillna("000.000.000-00", inplace = True)

EMISSOES['PA_BUC'] = np.where(EMISSOES['Vendedor'].str.contains('GESTAO ONLINE EXEMPLO', na= False),"GESTAO ONLINE EXEMPLO SHS",EMISSOES['PA_BUC'])
EMISSOES['Tipo de empresa - Cadastro PA'] = np.where(EMISSOES['Vendedor'].str.contains('GESTAO ONLINE EXEMPLO', na= False),"TESTE",EMISSOES['Tipo de empresa - Cadastro PA'])
EMISSOES['CPF - AGR'] = np.where(EMISSOES['Vendedor'].str.contains('GESTAO ONLINE EXEMPLO', na= False),"000.000.000-00",EMISSOES['CPF - AGR'])

EMISSOES[['Data','Data de aprovação']] = EMISSOES[['Data','Data de aprovação']].apply(lambda x: pd.to_datetime(x, errors='ignore', format = '%m/%d/%Y'))
EMISSOES.sort_values(by = ['Data de aprovação','Data','Conc IDAR'], ascending = True, inplace = True, ignore_index = True)
EMISSOES['Vendedor'].replace("SISTEMA #","SISTEMA-", inplace=True, regex = True)
EMISSOES['REPASSE AE ou PE'] = EMISSOES['REPASSE AE ou PE'].replace(["\$","\$-"],"", regex = True)
EMISSOES['REPASSE AE ou PE LIQ'] = EMISSOES['REPASSE AE ou PE LIQ'].replace(["\$","\$-"],"", regex = True)
EMISSOES['REPASSE EFETIVO AE ou PE'] = EMISSOES['REPASSE EFETIVO AE ou PE'].replace(["\$","\$-"],"", regex = True)
EMISSOES['GE'] = EMISSOES['GE'].replace(["\$","\$-"],"", regex = True)
EMISSOES['REPASSE GE'] = EMISSOES['REPASSE GE'].replace(["\$","\$-"],"", regex = True)
EMISSOES['CBO'] = EMISSOES['CBO'].replace(["\$","\$-"],"", regex = True)
EMISSOES['CNPJ PA_PROC'] = EMISSOES['CNPJ PA_PROC'].replace("nan","", regex = True)
EMISSOES['REPASSE GE'] = EMISSOES['REPASSE GE'].replace(["\$","\$-"],"", regex = True)

EMISSOES['CUSTO CENTRAL DE EMISSÃO'].replace("^(CADASTRAR)","", inplace=True, regex = True)
EMISSOES['Critério Remuneração CBO'].replace("^(CADASTRAR)","", inplace=True, regex = True)
EMISSOES['%PI CDB'].replace("^(CADASTRAR)","", inplace=True, regex = True)

EMISSOES.drop_duplicates(subset = 'Conc IDAR', keep = 'last', inplace = True, ignore_index = True)
EMISSOES.drop(['Valida_cobrança','TIPO DE PONTO'], axis=1, inplace=True)
EMISSOES = EMISSOES[['Identificador','Data','Data de aprovação','Situação','Vendedor','Cliente','E-mail','Telefone',
                     'Indicação','Valor total','Valor Total Nota','Valor Total Delivery','Itens do pedido de venda',
                     'Formas de pagamento do pedido de venda','Validação de Videoconferência','AR','Renovação','Renovado',
                     'Cliente Novo','A quem cobrar?','TABELA','PREÇO VENDA','TIPO','PERIODO DE COBRANÇA','Código AE ou PE',
                     'AE ou PE','% AE ou PE','REPASSE AE ou PE','REPASSE AE ou PE LIQ','REPASSE EFETIVO AE ou PE','GE','% GE',
                     'REPASSE GE','CUSTO(PE, AE e GE)','SITUAÇÃO DE PAGAMENTO','DESPESA BOLETO','DESPESA IMPOSTOS','DATA DINAMICA',
                     'DATA DINAMICA COMISSÃO','NOME FAIXA',' ABERTO','TIPO DE PARCEIRO','CUSTO ULT FAIXA','CBO','LIQUIDO',
                     'Status Soluti','Conc Midias','% PE 2','CUSTO CENTRAL DE EMISSÃO', 'RESULTADO NOSSO CERTIFICADO',
                     'CUSTO PARCEIRO INDICADDO','REPASSE PARCEIRO INDICADOR','CUSTAS PARCEIRO INDICADO','REPASSE PARCEIRO INDICADO',
                     'Franquia NTW','Validade','Tempo de Validade','Já venceu?','Tipo de Produto','Critério Remuneração CBO',
                     'Conc IDAR','Data de recebimento','procx','CICLO','REF','%PI CDB','Repasse PI CDB','Agente Captador',
                     'Repasse Agente Captador','PA_BUC','CNPJ PA','Tipo de empresa - Cadastro PA','CPF - AGR','CNPJ PA_PROC']]
EMISSOES.to_csv(f'{pasta_csv_buc}\\BUC\\EMISSOES.csv', sep=",", decimal='.', date_format='%m/%d/%Y %H:%M:%S', index=False, encoding='UTF-8')

#================================================RENOVAÇÕES===================================================

Renovacoes = EMISSOES[['Data de aprovação','Cliente','Itens do pedido de venda','Validade','Tempo de Validade','Vendedor','PA_BUC','CNPJ PA','Tipo de empresa - Cadastro PA','Conc IDAR']].copy()
Renovacoes['Validade'] = pd.to_datetime(Renovacoes['Validade'], errors='ignore', format = '%m/%d/%Y')
Renovacoes.rename(columns={'Vendedor':'Vendedor_Aprovação','PA_BUC':'PA_Aprovação','CNPJ PA':'CNPJ do PA_Aprovação','Tipo de empresa - Cadastro PA':'Tipo de Empresa_Aprovação'}, inplace = True)
Renovacoes['Cliente2'] = Renovacoes['Cliente']
Renovacoes['Cliente2'].replace("\(", " (", inplace = True, regex = True)
Renovacoes['Cliente2'].replace("\)", " )", inplace = True, regex = True)
Renovacoes['Cliente2'] = Renovacoes['Cliente2'].str.strip()
Renovacoes[['Nome do Cliente','Documento do Cliente2']] = Renovacoes['Cliente2'].str.rsplit('(', n= 1, expand = True)
Renovacoes['Documento do Cliente'] = Renovacoes['Documento do Cliente2'].replace(["\)","\.","\/","-"],"", regex = True)
Renovacoes['Documento do Cliente2'].replace("\)","", inplace = True, regex = True)

duplicados = Renovacoes[Renovacoes['Documento do Cliente'].duplicated()][['Nome do Cliente','Documento do Cliente']].copy()
duplicados['Documento do Cliente'].fillna("0", inplace = True)
duplicados['Documento do Cliente'] = duplicados['Documento do Cliente'].replace('nan',"0", regex = True)
duplicados = duplicados[(duplicados['Documento do Cliente'] != "0")].reset_index(drop = True)
duplicados.drop_duplicates(subset = 'Documento do Cliente', keep = 'first', inplace = True, ignore_index = True)
lista_duplicados = [duplicados.loc[i,'Documento do Cliente'] for i in range(len(duplicados['Documento do Cliente']))]
Renovacoes = Renovacoes[(Renovacoes['Documento do Cliente'].isin(lista_duplicados))]
Renovacoes.sort_values(by = ['Documento do Cliente','Data de aprovação'], ascending = True, inplace = True, ignore_index = True)
del duplicados
n_lin_ren = Renovacoes.shape[0]
valor_=0
for i in range(n_lin_ren):
    valor_=Renovacoes.loc[i,'Documento do Cliente']
    if i==n_lin_ren-1:
        Renovacoes.loc[i,'∆ Aprov_Renov'] = 0
    elif valor_ == Renovacoes.loc[i+1,'Documento do Cliente']:
        Renovacoes.loc[i,'ID renovação'] = Renovacoes.loc[i+1,'Conc IDAR']
        Renovacoes.loc[i,'Data renovação'] = Renovacoes.loc[i+1,'Data de aprovação']
        Renovacoes.loc[i,'Item renovação'] = Renovacoes.loc[i+1,'Itens do pedido de venda']
        Renovacoes.loc[i,'Vendedor_Renovação'] = Renovacoes.loc[i+1,'Vendedor_Aprovação']
        Renovacoes.loc[i,'PA_Renovação'] = Renovacoes.loc[i+1,'PA_Aprovação']
        Renovacoes.loc[i,'CNPJ do PA_Renovação'] = Renovacoes.loc[i+1,'CNPJ do PA_Aprovação']
        Renovacoes.loc[i,'Tipo de Empresa_Renovação'] = Renovacoes.loc[i+1,'Tipo de Empresa_Aprovação']        
        Renovacoes.loc[i,'∆ Aprov_Renov'] = Renovacoes.loc[i+1,'Data de aprovação'] - Renovacoes.loc[i,'Data de aprovação']
    else:
        Renovacoes.loc[i,'∆ Aprov_Renov'] = 0 
Renovacoes = Renovacoes[~(Renovacoes['Data renovação'].isna())][['Conc IDAR','Data de aprovação','Nome do Cliente','Documento do Cliente2','Vendedor_Aprovação','PA_Aprovação','CNPJ do PA_Aprovação',
                                                                 'Tipo de Empresa_Aprovação','Itens do pedido de venda','Validade','Tempo de Validade','ID renovação','Data renovação','Vendedor_Renovação',
                                                                 'PA_Renovação','CNPJ do PA_Renovação','Tipo de Empresa_Renovação','Item renovação','∆ Aprov_Renov']]
Renovacoes['Dif'] = Renovacoes['Data renovação'] - Renovacoes['Validade']
Renovacoes['Dif'] = Renovacoes['Dif'].astype(str)
Renovacoes['Dif'] = Renovacoes['Dif'].str.split(" ",expand=True)[0]
Renovacoes['∆ Aprov_Renov'] = Renovacoes['∆ Aprov_Renov'].astype(str)
Renovacoes['∆ Aprov_Renov'] = Renovacoes['∆ Aprov_Renov'].str.split(" ",expand=True)[0]
Renovacoes[['Tempo de Validade','Dif','∆ Aprov_Renov']] = Renovacoes[['Tempo de Validade','Dif','∆ Aprov_Renov']].apply(lambda x: x.astype(np.int64))
Renovacoes =  Renovacoes[((Renovacoes['Dif']>=-60) & (Renovacoes['Dif']<=360))].reset_index(drop = True)
Renovacoes.rename(columns={'Documento do Cliente2':'Documento do Cliente','Itens do pedido de venda':'Item aprovação','Conc IDAR':'ID aprovação','PA_BUC':'PA','CNPJ PA':'CNPJ do PA','Tipo de empresa - Cadastro PA':'Tipo de Empresa'}, inplace = True)
Renovacoes.drop('Dif', axis=1, inplace=True)
Renovacoes.sort_values(by = ['Data renovação','∆ Aprov_Renov'], ascending = [False,False], inplace = True, ignore_index = True)

ontem_ = dt.datetime.today()-dt.timedelta(1)
ontem_ = ontem_.strftime("%Y-%m-%d")
filtro_ontem = dt.datetime.fromisoformat(ontem_)
texto_ontem = dt.datetime.strptime(ontem_,"%Y-%m-%d").strftime("%d/%m/%Y")
renovados_ontem = Renovacoes[Renovacoes['Data renovação']==filtro_ontem][['ID aprovação','Data de aprovação','Nome do Cliente','Documento do Cliente','Item aprovação','Validade','Tempo de Validade','ID renovação','Data renovação','Item renovação','∆ Aprov_Renov']]
renovados_ontem.rename(columns={'∆ Aprov_Renov':'Intervalo de dias entre Aprovação e Renovação'}, inplace = True)
emissoes_ontem = EMISSOES[EMISSOES['Data de aprovação']==filtro_ontem]
if emissoes_ontem.shape[0]>0:
    result_ren = f'{(renovados_ontem.shape[0]/emissoes_ontem.shape[0])*100:.2f}%'.replace('.',',')
if renovados_ontem.shape[0]>0:
    renovados_ontem.to_excel(f'{pasta_csv_buc}\\Aviso Teams\\Renovações.xlsx', sheet_name = 'Renovações', index=False)

atualizacao = dt.datetime.now()
atualizacao1 = atualizacao.strftime("às %H:%M:%S do dia %d/%m/%Y")    

link_renovacoes = "[Ver Renovações]({link sharepoint})"
link_dashboard = "[Ver Dashboard]({link dashboard})"
time.sleep(5)

card_teams3 = teams.cardsection()
card_teams3.title(">NOVAS EMISSÕES")
card_teams4 = teams.cardsection()
card_teams4.title(">RENOVAÇÕES")

if emissoes_ontem.shape[0]>0:
    card_teams3.activityText(str(f"Tivemos {emissoes_ontem.shape[0]} emissões no dia {texto_ontem}. \n\n")+str("\n\n")+link_dashboard if emissoes_ontem.shape[0]>1 else str(f"Tivemos {emissoes_ontem.shape[0]} emissão no dia {texto_ontem}. \n\n")+str("\n\n")+link_dashboard)
else:
    card_teams3.activityText(str(f"""\nNão tivemos emissões no dia {texto_ontem}. \n\n""")+str("\n\n")+link_dashboard)

if renovados_ontem.shape[0]>0:
    card_teams4.activityText(str(f"Tivemos {renovados_ontem.shape[0]} renovações no dia {texto_ontem}. Isso representa {result_ren} das {emissoes_ontem.shape[0]} emissões no dia {texto_ontem}. \n\n")+str("\n\n")+link_renovacoes if renovados_ontem.shape[0]>1 else str(f"Tivemos {renovados_ontem.shape[0]} renovação no dia {texto_ontem}. Isso representa {result_ren} das {emissoes_ontem.shape[0]} emissões no dia {texto_ontem}. \n\n")+str("\n\n")+link_renovacoes)
else:
    card_teams4.activityText(str(f"""\nNão tivemos renovações no dia {texto_ontem}. \n\n"""))
card_teams4.text(str("\n\n")+str(f'### Atualizado {atualizacao1}.\n\n')+str("\n\n")+str("""## OBS: As Dashboards são atualizadas de segunda a sexta, às 14:40. \n\n"""))


msg_teams1.addSection(card_teams3)
msg_teams1.addSection(card_teams4)
msg_teams1.send()
Renovacoes.to_csv(f'{pasta_csv_buc}\\Renovacoes.csv', sep=",", decimal='.', date_format='%d/%m/%Y %H:%M:%S', index=False, encoding='UTF-8')

#================================================CERTIFICADOS VINCENDOS===================================================

hoje = dt.date.today()
oito_dias = hoje-dt.timedelta(8)
oito_dias = oito_dias.strftime("%Y-%m-%d")
filtro_vincendos = hoje + dt.timedelta(days=90)
filtro_data_oito_dias = dt.datetime.fromisoformat(oito_dias)
filtro_data_vincendos = f'{filtro_vincendos.year}-{filtro_vincendos.month}-{filtro_vincendos.day}'
Cert_vincendos = EMISSOES[['Conc IDAR','Cliente','E-mail','Telefone','Itens do pedido de venda','Data de aprovação','Validade','Vendedor','PA_BUC']].copy()
Cert_vincendos['Validade'] = pd.to_datetime(Cert_vincendos['Validade'], errors='ignore', format = '%m/%d/%Y')
Cert_vincendos = Cert_vincendos[((Cert_vincendos['Validade']>filtro_data_oito_dias) & (Cert_vincendos['Validade']<=filtro_data_vincendos))].reset_index( drop = True)
Cert_vincendos['Dias até o vencimento'] = Cert_vincendos['Validade']-pd.to_datetime(hoje)
Cert_vincendos['Dias até o vencimento'] = Cert_vincendos['Dias até o vencimento'].astype(str)
Cert_vincendos['Dias até o vencimento'] = Cert_vincendos['Dias até o vencimento'].str.split(" ",expand=True)[0]
Cert_vincendos['Dias até o vencimento'] = Cert_vincendos['Dias até o vencimento'].astype(np.int64)
Cert_vincendos.sort_values(by = ['Validade','Dias até o vencimento'], ascending = True, inplace = True, ignore_index = True)
Cert_vincendos.to_csv(f'{pasta_csv_buc}\\Cert_vincendos.csv', sep=",", decimal='.', date_format='%d/%m/%Y %H:%M:%S', index=False, encoding='UTF-8')

#================================================EMISSÃO MÊS A MÊS POR PA===================================================

EMISSOES_MES_A_MES = EMISSOES[['Conc IDAR','Data de aprovação','CNPJ PA_PROC']].copy()
EMISSOES_MES_A_MES = EMISSOES_MES_A_MES[~((EMISSOES_MES_A_MES['CNPJ PA_PROC'].isna()) | (EMISSOES_MES_A_MES['CNPJ PA_PROC']=='') | (EMISSOES_MES_A_MES['CNPJ PA_PROC']=='nan'))]
EMISSOES_MES_A_MES.sort_values(by=['CNPJ PA_PROC','Data de aprovação'], ascending = True, inplace = True, ignore_index = True)
EMISSOES_MES_A_MES['Data de aprovação'] = EMISSOES_MES_A_MES['Data de aprovação'].dt.strftime("%m/%Y")
EMISSOES_MES_A_MES.rename(columns={'Conc IDAR':'Qtd_Emissões'}, inplace = True)
EMISSOES_MES_A_MES = EMISSOES_MES_A_MES.groupby(by=['Data de aprovação','CNPJ PA_PROC'], as_index = False, dropna = False, sort = False).count()[['Data de aprovação','CNPJ PA_PROC','Qtd_Emissões']]
EMISSOES_MES_A_MES = EMISSOES_MES_A_MES.join(CADASTRO_PA[['CNPJ_PROC','DOCUMENTO PA','Nome da Empresa']].set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['Data de aprovação','CNPJ PA_PROC','DOCUMENTO PA','Nome da Empresa','Qtd_Emissões']]
EMISSOES_MES_A_MES.to_csv(f'{pasta_csv_buc}\\EMISSOES_MES_A_MES.csv', sep=",", decimal='.', date_format='%d/%m/%Y', index=False, encoding='UTF-8')

#================================================EMISSÃO MÊS A MÊS POR AE ou PE===================================================

EMISSOES_MES_A_MES_AE_PE = EMISSOES[['Conc IDAR','Data de aprovação','AE ou PE']].copy()
EMISSOES_MES_A_MES_AE_PE = EMISSOES_MES_A_MES_AE_PE[~((EMISSOES_MES_A_MES_AE_PE['AE ou PE'].isna()) | (EMISSOES_MES_A_MES_AE_PE['AE ou PE']=='') | (EMISSOES_MES_A_MES_AE_PE['AE ou PE']=='nan'))]
EMISSOES_MES_A_MES_AE_PE.sort_values(by=['AE ou PE','Data de aprovação'], ascending = True, inplace = True, ignore_index = True)
EMISSOES_MES_A_MES_AE_PE['Data de aprovação'] = EMISSOES_MES_A_MES_AE_PE['Data de aprovação'].dt.strftime("%m/%Y")
EMISSOES_MES_A_MES_AE_PE.rename(columns={'Conc IDAR':'Qtd_Emissões'}, inplace = True)
EMISSOES_MES_A_MES_AE_PE = EMISSOES_MES_A_MES_AE_PE.groupby(by=['Data de aprovação','AE ou PE'], as_index = False, dropna = False, sort = False).count()[['Data de aprovação','AE ou PE','Qtd_Emissões']]
EMISSOES_MES_A_MES_AE_PE.to_csv(f'{pasta_csv_buc}\\EMISSOES_MES_A_MES_AE_PE.csv', sep=",", decimal='.', date_format='%d/%m/%Y', index=False, encoding='UTF-8')

#================================================EMISSÕES 15,30,45,60,90,180 DIAS POR PA===================================================

filtro_365dias = dt.date.today() - dt.timedelta(days=366)
filtro_data_365dias = f'{filtro_365dias.year}-{filtro_365dias.month}-{filtro_365dias.day}'
tit_365dias = f'Emissões 365 dias:{filtro_365dias.strftime("%d/%m/%Y")}'

filtro_180dias = dt.date.today() - dt.timedelta(days=181)
filtro_data_180dias = f'{filtro_180dias.year}-{filtro_180dias.month}-{filtro_180dias.day}'
tit_180dias = f'Emissões 180 dias:{filtro_180dias.strftime("%d/%m/%Y")}'

filtro_90dias = dt.date.today() - dt.timedelta(days=91)
filtro_data_90dias = f'{filtro_90dias.year}-{filtro_90dias.month}-{filtro_90dias.day}'
tit_90dias = f'Emissões 90 dias:{filtro_90dias.strftime("%d/%m/%Y")}'

filtro_60dias = dt.date.today() - dt.timedelta(days=61)
filtro_data_60dias = f'{filtro_60dias.year}-{filtro_60dias.month}-{filtro_60dias.day}'
tit_60dias = f'Emissões 60 dias:{filtro_60dias.strftime("%d/%m/%Y")}'

filtro_45dias = dt.date.today() - dt.timedelta(days=46)
filtro_data_45dias = f'{filtro_45dias.year}-{filtro_45dias.month}-{filtro_45dias.day}'
tit_45dias = f'Emissões 45 dias:{filtro_45dias.strftime("%d/%m/%Y")}'

filtro_30dias = dt.date.today() - dt.timedelta(days=31)
filtro_data_30dias = f'{filtro_30dias.year}-{filtro_30dias.month}-{filtro_30dias.day}'
tit_30dias = f'Emissões 30 dias:{filtro_30dias.strftime("%d/%m/%Y")}'

filtro_15dias = dt.date.today() - dt.timedelta(days=16)
filtro_data_15dias = f'{filtro_15dias.year}-{filtro_15dias.month}-{filtro_15dias.day}'
tit_15dias = f'Emissões 15 dias:{filtro_15dias.strftime("%d/%m/%Y")}'

EMITIU_ESSE_MES = EMISSOES[((EMISSOES['Data de aprovação']>=filtro_inicio_mes) & (EMISSOES['Data de aprovação']<=filtro_final_mes))].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMITIU_ESSE_MES.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMITIU_ESSE_MES.rename(columns={'Conc IDAR':'Emissões nesse mês'}, inplace = True)

EMITIU_MES_ANT = EMISSOES[((EMISSOES['Data de aprovação']>=filtro_inicio_mes_ant) & (EMISSOES['Data de aprovação']<=filtro_final_mes_ant))].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMITIU_MES_ANT.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMITIU_MES_ANT.rename(columns={'Conc IDAR':'Emissões mês anterior'}, inplace = True)

EMISSOES_por_pa = EMISSOES.groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_por_pa = EMISSOES_por_pa[EMISSOES_por_pa['CNPJ PA_PROC'].notna()].reset_index(drop=True)
EMISSOES_por_pa.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_por_pa.rename(columns={'Conc IDAR':'Emissões por PA'}, inplace = True)

EMISSOES_365dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_365dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_365dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_365dias.rename(columns={'Conc IDAR':tit_365dias}, inplace = True)

EMISSOES_180dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_180dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_180dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_180dias.rename(columns={'Conc IDAR':tit_180dias}, inplace = True)

EMISSOES_90dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_90dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_90dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_90dias.rename(columns={'Conc IDAR':tit_90dias}, inplace = True)

EMISSOES_60dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_60dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_60dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_60dias.rename(columns={'Conc IDAR':tit_60dias}, inplace = True)

EMISSOES_45dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_45dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_45dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_45dias.rename(columns={'Conc IDAR':tit_45dias}, inplace = True)

EMISSOES_30dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_30dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_30dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_30dias.rename(columns={'Conc IDAR':tit_30dias}, inplace = True)

EMISSOES_15dias = EMISSOES[EMISSOES['Data de aprovação']>=filtro_data_15dias].groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
EMISSOES_15dias.sort_values('Conc IDAR', ascending = False, inplace = True, ignore_index = True)
EMISSOES_15dias.rename(columns={'Conc IDAR':tit_15dias}, inplace = True)

CADASTRO_PA = CADASTRO_PA.join(EMITIU_ESSE_MES.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMITIU_MES_ANT.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_15dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_30dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_45dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_60dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_90dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_180dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_365dias.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')
CADASTRO_PA = CADASTRO_PA.join(EMISSOES_por_pa.set_index('CNPJ PA_PROC'), on = 'CNPJ_PROC', how = 'left',rsuffix='_')

col_emissoes = [i for i in CADASTRO_PA.columns if 'Emissões' in i]
CADASTRO_PA[col_emissoes] = CADASTRO_PA[col_emissoes].apply(lambda x: x.replace(np.nan,0))
CADASTRO_PA[col_emissoes] = CADASTRO_PA[col_emissoes].apply(lambda x: x.astype(np.int64))
CADASTRO_PA.sort_values(by='Criado', ascending = False, inplace = True, ignore_index = True)

CADASTRO_PA.to_csv(f'{pasta_csv_buc}\\CADASTRO_PA.csv', sep=",", decimal = ".", date_format = '%d/%m/%Y %H:%M:%S', index=False, encoding='UTF-8')

#=============================================RANKING EMISSÕES SHS==============================================

mes_ant_ = f'Emissões Mês {filtro_inicio_mes_ant.strftime("%m/%Y")}'
mes_atual_ = f'Emissões Mês {filtro_inicio_mes.strftime("%m/%Y")}'
mes_ant_1 = f'Mês {filtro_inicio_mes_ant.strftime("%m/%Y")}'
mes_atual_1 = f'Mês {filtro_inicio_mes.strftime("%m/%Y")}'

filtro_shs = ['SHS PI','SHS PV']
dt_shs = EMISSOES[(EMISSOES['Tipo de empresa - Cadastro PA'].isin(filtro_shs))][['CNPJ PA_PROC','CNPJ PA','Data de aprovação','Conc IDAR']].copy().reset_index(drop=True)
Emissoes_SHS = EMISSOES[(EMISSOES['Tipo de empresa - Cadastro PA'].isin(filtro_shs))][['CNPJ PA_PROC','Conc IDAR']].copy().reset_index(drop=True)
Emissoes_SHS = Emissoes_SHS.groupby(by='CNPJ PA_PROC',as_index=False,dropna=False).count()[['CNPJ PA_PROC','Conc IDAR']]
n_emi = sorted(list(set(Emissoes_SHS['Conc IDAR'])),reverse=True)
Emissoes_SHS = Emissoes_SHS[Emissoes_SHS['Conc IDAR'].isin(n_emi[:3])].reset_index(drop = True).sort_values('Conc IDAR', ascending = False, ignore_index = True)
for i in range(len(Emissoes_SHS)):
    Emissoes_SHS.loc[i,'Ranking'] = f"{n_emi.index(Emissoes_SHS.loc[i,'Conc IDAR'])+1}º"
    Emissoes_SHS.loc[i,'Última Emissão'] = max(dt_shs[dt_shs['CNPJ PA_PROC']==Emissoes_SHS.loc[i,'CNPJ PA_PROC']]['Data de aprovação'])
Emissoes_SHS['Última Emissão'] = Emissoes_SHS['Última Emissão'].dt.strftime("%d/%m/%Y")
Emissoes_SHS = Emissoes_SHS.join(CADASTRO_PA[['CNPJ_PROC','DOCUMENTO PA','Nome da Empresa','Emissões nesse mês','Emissões mês anterior']].set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['Ranking','DOCUMENTO PA','Nome da Empresa','Conc IDAR','Emissões nesse mês','Emissões mês anterior','Última Emissão']]
Emissoes_SHS.rename(columns={'CNPJ PA_PROC':'CNPJ','Conc IDAR':'Total de Emissões','DOCUMENTO PA':'CNPJ','Nome da Empresa':'Software House','Emissões nesse mês':mes_atual_,'Emissões mês anterior':mes_ant_}, inplace = True)
del dt_shs

tabulate.PRESERVE_WHITESPACE=True
tabela = [['Ranking','Software House','Total', mes_atual_1, mes_ant_1,'Última Emissão']]
linha = []
for i in range(len(Emissoes_SHS)):
    col_ranking = Emissoes_SHS.loc[i,'Ranking']
    col_shs = f"{Emissoes_SHS.loc[i,'Software House']:^6}"
    col_qtd = f"{Emissoes_SHS.loc[i,'Total de Emissões']:^17}"
    col_emi_mes = f"{Emissoes_SHS.loc[i, mes_atual_]:^17}"
    col_emi_ant = Emissoes_SHS.loc[i, mes_ant_]
    col_ult_emi = Emissoes_SHS.loc[i,'Última Emissão']
    linha = [col_ranking,col_sh,col_qtd,col_emi_mes,col_emi_ant,col_ult_emi]
    tabela.append(linha)
tabela_formatada = tabulate.tabulate(tabela,headers="firstrow",tablefmt="grid",colalign=('center','left','center','center','center'),maxcolwidths=[10,None,15,15,15])

#=============================================RANKING EMISSÕES POR PA==============================================

n_emi_pas = sorted(list(set(EMISSOES_por_pa['Emissões por PA'])),reverse=True)[:10]
Analise_PA = EMISSOES_por_pa[EMISSOES_por_pa['Emissões por PA'].isin(n_emi_pas)].reset_index(drop = True).sort_values('Emissões por PA', ascending = False, ignore_index = True)
dt_pas = EMISSOES[(EMISSOES['CNPJ PA_PROC'].isin(list(EMISSOES_por_pa['CNPJ PA_PROC'])))][['CNPJ PA_PROC','Data de aprovação','Conc IDAR']].copy().reset_index(drop=True)
for i in range(len(Analise_PA)):
    Analise_PA.loc[i,'Ranking'] = f"{n_emi_pas.index(Analise_PA.loc[i,'Emissões por PA'])+1}º"
    Analise_PA.loc[i,'Última Emissão'] = max(dt_pas[dt_pas['CNPJ PA_PROC']==Analise_PA.loc[i,'CNPJ PA_PROC']]['Data de aprovação'])
Analise_PA = Analise_PA.join(RUN_AGR.drop_duplicates(subset = 'CNPJ_PROC', keep = 'last')[['CNPJ_PROC','DOCUMENTO PA','PA','Tipo de Ponto']].set_index('CNPJ_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')
Analise_PA = Analise_PA.join(EMITIU_ESSE_MES.set_index('CNPJ PA_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')
Analise_PA = Analise_PA.join(EMITIU_MES_ANT.set_index('CNPJ PA_PROC'), on = 'CNPJ PA_PROC', how = 'left',rsuffix='_')[['Ranking','DOCUMENTO PA','PA','Emissões por PA','Emissões nesse mês','Emissões mês anterior','Última Emissão']]
Analise_PA['Última Emissão'] = Analise_PA['Última Emissão'].dt.strftime("%d/%m/%Y")
col_emissoes = [i for i in Analise_PA.columns if 'Emissões' in i]
Analise_PA[col_emissoes] = Analise_PA[col_emissoes].apply(lambda x: x.replace(np.nan,0))
Analise_PA[col_emissoes] = Analise_PA[col_emissoes].apply(lambda x: x.astype(np.int64))
Analise_PA.rename(columns={'DOCUMENTO PA':'CNPJ/CPF','Emissões por PA':'Total de Emissões','Emissões nesse mês':mes_atual_,'Emissões mês anterior':mes_ant_}, inplace = True)
if '00.000.000/0001-00' in list(Analise_PA['CNPJ/CPF']):
    Analise_PA['PA'] = np.where(Analise_PA['CNPJ/CPF']=='00.000.000/0001-00',"VENDAS INTERNAS",Analise_PA['PA'])
del dt_pas

tabulate.PRESERVE_WHITESPACE=True
tabela1 = [['Ranking','PA','Total', mes_atual_1, mes_ant_1,'Última Emissão']]
linha = []
for i in range(len(Analise_PA)):
    col_ranking = Analise_PA.loc[i,'Ranking']
    col_sh = f"{Analise_PA.loc[i,'PA']:^6}"
    col_qtd = f"{Analise_PA.loc[i,'Total de Emissões']:^17}"
    col_emi_mes = f"{Analise_PA.loc[i, mes_atual_]:^17}"
    col_emi_ant = Analise_PA.loc[i, mes_ant_]
    col_ult_emi = Analise_PA.loc[i,'Última Emissão']
    linha = [col_ranking,col_sh,col_qtd,col_emi_mes,col_emi_ant,col_ult_emi]
    tabela1.append(linha)
tabela_formatada1 = tabulate.tabulate(tabela1,headers="firstrow",tablefmt="grid",colalign=('center','left','center','center','center'),maxcolwidths=[10,None,15,15,15])

card_teams5 = teams.cardsection()
card_teams5.title(">SHS")
card_teams5.activityText(str('## → Top 3 SHS. \n\n')+tabela_formatada)
msg_teams2.addSection(card_teams5)

card_teams6 = teams.cardsection()
card_teams6.title(">EMISSÕES POR PA")
card_teams6.activityText(str('## → Top 10 PAs. \n\n')+tabela_formatada1)
msg_teams2.addSection(card_teams6)
# msg_teams2.send()

#================================================ATUALIZAÇÃO DAS BASES===================================================

app1 = xw.App(visible=True, add_book=False)
app1.display_alerts=False
wb1 = app1.books.open(f'{pasta_base}BUC - Cadastros.xlsx')
titulo1 = 'BUC - Cadastros.xlsx - Excel'
a1 =gw.getWindowsWithTitle(titulo1)[0]
time.sleep(5)
a1.maximize()
app1.calculation = 'manual'
wb1.api.RefreshAll()
time.sleep(300) # 5 minutos
app1.calculation = 'automatic'
time.sleep(30)
wb1.save()
time.sleep(10)
wb1.close()
time.sleep(10)

wb2 = app1.books.open(f'{pasta_base}BUC - Emissões.xlsx')
titulo2 = 'BUC - Emissões.xlsx - Excel'
a2 =gw.getWindowsWithTitle(titulo2)[0]
time.sleep(5)
a2.maximize()
app1.calculation = 'manual'
wb2.api.RefreshAll()
time.sleep(420) # 7 minutos
app1.calculation = 'automatic'
time.sleep(30)
wb2.save()
time.sleep(10)
wb2.close()
time.sleep(10)
app1.quit()