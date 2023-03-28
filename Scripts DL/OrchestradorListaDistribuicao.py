from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.keys import Keys
import pydirectinput
import keyboard
import os, sys
import pyodbc

#import GERAL_LoginClick as LogarNoClick
import GERALFiltrarDemandaClick as AplicaFiltro
import LISTADISTRIBUICAO_PegaDemanda as PegaDemanda
import LISTADISTRIBUICAO_CriaLista as CriaListas
import GERAL_FechaChamadoClick as FechaChamado

import time
import datetime

def LerTabela():
    try:
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        cur = con.cursor()
        cur.execute("Select * from ListaDistribuicao where StatusCriaLista = 'OK' AND StatusFechaChamado = 'PENDENTE'")
        
        dbData = cur.fetchall()
        if (len(dbData) > 0):
            WnumeroChamado = str(dbData[0][1])
            WNomeLista = str(dbData[0][15])
            WOpcao = str(dbData[0][2])
        else:
            WnumeroChamado = '00000000'
            WNomeLista = ''
            WOpcao = ''
         
        cur.close()
        con.close()

        return WnumeroChamado, WNomeLista, WOpcao

    except Exception as e:
        print('****ERRO****:', e)

first_time = datetime.datetime.now()

browser = profile = webdriver.FirefoxProfile()
browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))

try:
    browser.get("http://gestaoclick.petrobras.com.br/")

    time.sleep(3)
    # Faz o LOGIN na tela do Navegador
    pydirectinput.moveTo(790, 200)
    pydirectinput.click()
    keyboard.write('SAMSAZU@petrobras.com.br')
    time.sleep(1)
    pydirectinput.press('tab')
    time.sleep(1)
    keyboard.write('Automacao01@') 
    time.sleep(1)
    pydirectinput.press('tab')
    time.sleep(1)
    pydirectinput.press('return')

    element = WebDriverWait(browser,60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contentPaneTitle"]')))
except Exception as e:
    print('****ERRO NO ACESSO A URL DO CLICK****', e) 

# Executa a rotina de Login no site do GestaoClick
#ExecutaLogin = LogarNoClick.Login('SAMSAZU@petrobras.com.br', browser)
#if (ExecutaLogin != 'OK'):
#    print('Houve ERRO na rotina de Login')
#    sys.exit()

#---------------------------------------------------------------------------------------------------------------#
# Executa a rotina de aplicação de Filtros
print('** Inicio AplicaFiltro')
ExecutaFiltros = AplicaFiltro.Filtrar('Criação/Alteração/Exclusão de lista de distribuição - Outlook', browser)
if (ExecutaFiltros != 'OK'):
    print('Houve ERRO na rotina de Filtrar solicitações', ExecutaFiltros)
    sys.exit()
print('** Fim AplicaFiltro')

#---------------------------------------------------------------------------------------------------------------#
# Executa a rotina para pegar os dados Automacao01@ 
# e salvar no banco
print('** Inicio PegaDemanda')
SalvarDados = PegaDemanda.SalvaDemanda(browser)
if (SalvarDados == 'Sem Chamados'):
    print('*-------------------------------------*')
    print('* Não existem chamados para processar *')
    print('*-------------------------------------*')
    browser.close()

if (SalvarDados != 'OK' and SalvarDados != 'Sem Chamados'):
    print('Houve ERRO na rotina de Salvar no Banco', SalvarDados)
    sys.exit()
print('** Fim PegaDemanda')

#---------------------------------------------------------------------------------------------------------------#
# Executa a rotina de Manutenção da Lista de Distribuição
print('** Inicio CriaLista')

CriaLista = CriaListas.RotinaPrincipal()
if (CriaLista == 'OK'):
    print('*-------------------------------------*')
    print('* A Lista foi criada com sucesso      *')
    print('*-------------------------------------*')

print('** Fim CriaLista')

#---------------------------------------------------------------------------------------------------------------#
# Executa a rotina de Fechamento do Chamado
# pega os chamados em aberto
print('** Inicio FechaChamado')
NomeLista = ''
NumeroChamado, NomeLista, Opcao = LerTabela()

if (Opcao == 'Criar (menos que 100 membros)'):
    DescricaoFechamento = 'Conforme solicitado, a Lista de Distribuição com o nome: <DL> e E-mail: <E-MAILDL> foi criada e os usuários indicados foram adicionados.'
if (Opcao == 'Excluir lista'):
    DescricaoFechamento = 'Conforme solicitado, a Lista de Distribuição com o nome: <DL> e E-mail: <E-MAILDL> foi excluída.'
if (Opcao == 'Incluir chaves'):
    DescricaoFechamento = 'As chaves solicitadas foram incluídas na Lista de Distribuição com o nome: <DL> e E-mail: <E-MAILDL> '
if (Opcao == 'Excluir chaves'):
    DescricaoFechamento = 'As chaves solicitadas foram excluídas na Lista de Distribuição com o nome: <DL> e E-mail: <E-MAILDL> '

if (NumeroChamado != '00000000'):
    DescricaoFechamento = DescricaoFechamento.replace('<DL>',NomeLista)
    DescricaoFechamento = DescricaoFechamento.replace('<E-MAILDL>',NomeLista + '@petrobras.com.br')
    FechaCh = FechaChamado.FechaChamado(NumeroChamado, 'N3-MICROSOFT_OFFICE', DescricaoFechamento, 'Office 365', 'Outros')
else:
    FechaCh = ''

#Whandles = browser.window_handlesAutomacao01@  


if (FechaCh == 'OK'):
    print('*-------------------------------------------*')
    print('* O chamado ', NumeroChamado, ' foi fechado com sucesso     *')
    print('*-------------------------------------------*')
else:
    #print('*---------------------------------------*')
    #print('* Houve Erro no fechamento do chamado   *')
    print (FechaCh)
    #print('*---------------------------------------*')
    #browser.close()

print('** Fim FechaChamado')

#---------------------------------------------------------------------------------------------------------------#
later_time = datetime.datetime.now()
difference = later_time - first_time
datetime.timedelta(0, 8, 562000)
seconds_in_day = 24 * 60 * 60
duracao = divmod(difference.days * seconds_in_day + difference.seconds, 60)
print('*----------------------------------------------------------*')
print('* ------ Finalizado em ', duracao[0], 'Minutos e ', duracao[1], ' Segundos', '------ *')
print('*----------------------------------------------------------*')
    
a=1