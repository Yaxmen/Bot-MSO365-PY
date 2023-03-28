import time, sys
import subprocess
import re
import pyodbc
import linecache, sys
from  datetime import datetime
from selenium import webdriver #import webdriver module
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By

def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print ('EXCEPTION IN (LINE {} "{}"): {}'.format(lineno, line.strip(), exc_obj))

def PegaEmailDaChave(Chave):
    # Chama o Script em PowerShell para pegar o Email relativo a Chave informada
    p=subprocess.Popen(['powershell.exe', 'D:\\Testes\\PegaEmail.ps1',Chave], stdout=subprocess.PIPE)

    retorno=p.communicate()[0]
    #print (retorno)
    formata = retorno.strip().decode("latin-1") 
    #formata = retorno.strip().decode("utf-8")

    # localiza os emails através de REGEX
    textoEmail = re.findall("[\w\-.]+@[\w\-]+\.\w+\.?\w*", formata)

    # Localiza o Nome completo do Usuário
    NomeCompleto = re.findall("[^\\n]*$", formata)

    # A segunda ocorrencia do Array tem o email relativo a chave pesquisada
    return textoEmail[1], NomeCompleto[0]

def MontaStringComEmails(Chaves):
    Arraychaves = Chaves.split(',')
    txtEmail = ''
    txtNome = ''
    number= 0
    
    while number < len(Arraychaves) :  
        retornoEmail = PegaEmailDaChave(Arraychaves[number])
        txtEmail = txtEmail + ' ' + retornoEmail[0]
        WajustaNome = retornoEmail[1].replace('\x87','ç')
        WajustaNome = WajustaNome.replace('£','ú')
        txtNome = txtNome + WajustaNome + '|'
        number = number+1

    txtNome = txtNome[:-1]
    return txtEmail, txtNome

def ChecaChamado(NumeroChamado):
    
    try:
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        cur = con.cursor()
        cur.execute("Select IdListaDistribuicao from ListaDistribuicao where NumeroChamado = " + "'" + NumeroChamado + "'")
        
        dbData = cur.fetchall()
        
        cur.close()
        con.close()

        return dbData
    except Exception as e:
        print('****CHECACHAMADO****:', e)
        
def Grava_dados_ListaDistribuicao(numeroChamadoStr, opcao, opcaoNome1, opcaoNome2, opcaoNome3, membros, proprietarios, remetentes, aberta, fechada, StringEmailMembros, StringEmailProprietarios, DataCriacao, NomeCompletoMembros, NomeCompletoProprietarios, emailDaLista):
    
    try:
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        mycursor = con.cursor()

        sql = "INSERT INTO dbo.ListaDistribuicao (NumeroChamado, Opcao, OpcaoNome1, OpcaoNome2, OpcaoNome3, ChavesMembros, ChavesProprietarios, ChavesRemetentes, OpcaoIngressar, OpcaoSair, StatusCriaLista, StatusFechaChamado, EmailMembros, EmailProprietarios, DataCriacao, NomeCompletoMembros, NomeCompletoProprietarios, EmailDaLista) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        val = (numeroChamadoStr, opcao,  opcaoNome1, opcaoNome2, opcaoNome3, membros, proprietarios, remetentes, aberta, fechada, 'PENDENTE', 'PENDENTE', StringEmailMembros, StringEmailProprietarios, DataCriacao, NomeCompletoMembros, NomeCompletoProprietarios, emailDaLista)

        mycursor.execute(sql, val)

    except Exception as erro:
        print('Houve um ERRO na gravação do Banco', erro)
        return "ERRO", erro, ""  

    con.commit()
    con.close()
    return 'OK'

def SalvaDemanda(browser):

    try:
        time.sleep(4)
        resultado = browser.find_element(By.XPATH,'//*[@id="eventMonitorCaption"]')
        textoResultado = resultado.text
        if 'Não foram encontrados resultados para os critérios especificados' in textoResultado:
            return 'Sem Chamados'
        
        #Abrindo chamado
        source = browser.find_element(By.XPATH,'/html/body/section[1]/div[2]/div/div/div/div[4]/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div[2]/div/div/div/div/div')
        time.sleep(3)
        action = ActionChains(browser)
        action.double_click(source)
        action.perform()

        element = WebDriverWait(browser,60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contentPaneTitle"]')))

        time.sleep(5)
        #Pegando numero do chamado
        c0 = browser.find_element(By.XPATH,'//*[@id="contentPaneTitle"]')
        time.sleep(3)
        numeroChamado = c0.text
        numeroChamadoStr = str(numeroChamado[0:8])
        #numeroChamadoStr = numeroChamadoStr.strip()
        #numeroChamadoStr = numeroChamadoStr.replace('(Aberto)','')
        #numeroChamadoStr = str(c0.text)

        print('====chamado====>',numeroChamadoStr)
        
        element1 = WebDriverWait(browser,60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="menuActions_label"]')))
        
        #Pegando info do campo opcao
        c1 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59080_textNode"]') #//*[@id="EC2791_webCustomPropertyValue_58620_textNode"] 
        opcao = c1.get_attribute('value')
        
        # o XPath muda dependendo da opção escolhida (Criação de Lista, Inclusão de chaves..etc)
        # Configurar o XPath dependendo da OPÇÃO
        if (opcao == 'Criar (menos que 100 membros)'):
            #Pegando info do campo opcaoNome1
            c2 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59081"]')  # 
            opcaoNome1 = c2.get_attribute('value')
            #Pegando info do campo opcaoNome2
            c3 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59082"]')  # 
            opcaoNome2 = c3.get_attribute('value')
            #Pegando info do campo opcaoNome3
            c4 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59083"]')  # 
            opcaoNome3 = c4.get_attribute('value')
            #Pegando info do campo membros
            c5 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59084"]') # 
            membros = c5.get_attribute('value')
            #Pegando info do campo proprietarios
            StringEmailMembros, txtNomeCompletoMembros = MontaStringComEmails(membros)
            c6 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59085"]') #
            proprietarios = c6.get_attribute('value')
            StringEmailProprietarios, txtNomeCompletoProprietarios = MontaStringComEmails(proprietarios)
            #Pegando info do campo Remetentes
            c61 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59101"]') # 
            remetentes = c61.get_attribute('value')
            #Pegando info do campo aberta
            c7 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59086_textNode"]') # 
            aberta = c7.get_attribute('value')
            #Pegando info do campo fechada
            c8 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59087_textNode"]') # 
            fechada = c8.get_attribute('value')
            emailDaLista = ''
        
        if (opcao == 'Incluir chaves'):
            #Pegando info do campo membros
            c5 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59088"]') #
            membros = c5.get_attribute('value')
            #Pegando info do campo proprietarios
            StringEmailMembros, txtNomeCompletoMembros = MontaStringComEmails(membros)
            c6 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59089"]') #
            proprietarios = c6.get_attribute('value')
            StringEmailProprietarios, txtNomeCompletoProprietarios = MontaStringComEmails(proprietarios)
            #Pegando info do campo Remetentes
            c61 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59108"]') # 
            remetentes = c61.get_attribute('value')            
            #Pegando info do campo aberta
            c7 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59090_textNode"]') # 
            aberta = c7.get_attribute('value')
            #Pegando info do campo fechada
            c8 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59091_textNode"]') # 
            fechada = c8.get_attribute('value')
            #Pegando Email da Lista
            c9 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59092"]') #
            emailDaLista = c9.get_attribute('value')
            opcaoNome1 = ''
            opcaoNome2 = ''
            opcaoNome3 = ''

        if (opcao == 'Excluir chaves'):
            #Pegando info do campo membros
            c5 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59093"]') # 
            membros = c5.get_attribute('value')
            #Pegando info do campo proprietarios
            StringEmailMembros, txtNomeCompletoMembros = MontaStringComEmails(membros)
            c6 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59094"]') #
            proprietarios = c6.get_attribute('value')
            StringEmailProprietarios, txtNomeCompletoProprietarios = MontaStringComEmails(proprietarios)
            #Pegando info do campo Remetentes
            c61 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59108"]') # 
            remetentes = c61.get_attribute('value')           
            #Pegando info do campo aberta
            #c7 = browser.find_element(By.XPATH,'//*[@id="EC2791_webCustomPropertyValue_58637_textNode"]') # 
            #aberta = c7.get_attribute('value')
            #Pegando info do campo fechada
            #c8 = browser.find_element(By.XPATH,'//*[@id="EC2791_webCustomPropertyValue_58638_textNode"]') # 
            #fechada = c8.get_attribute('value')
            #Pegando Email da Lista
            c9 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59097"]') # 
            emailDaLista = c9.get_attribute('value')
            opcaoNome1 = ''
            opcaoNome2 = ''
            opcaoNome3 = ''
            aberta = ''
            fechada = ''

        if (opcao == 'Excluir lista'):
            #Pegando Email da Lista
            c9 = browser.find_element(By.XPATH,'//*[@id="EC2818_webCustomPropertyValue_59099"]') #//*[@id="EC2791_webCustomPropertyValue_58630"] 
            emailDaLista = c9.get_attribute('value')
            opcaoNome1 = ''
            opcaoNome2 = ''
            opcaoNome3 = ''
            membros = ''
            proprietarios = ''
            remetentes = ''
            aberta = ''
            fechada = ''
            StringEmailMembros = ''
            StringEmailProprietarios = ''
            txtNomeCompletoMembros = ''
            txtNomeCompletoProprietarios = ''

        #print('===>  PEGADEMANDA LINHA ==> 109')
        
        # Verifica se o numero do chamado já existe no banco
        temChamado = ChecaChamado(numeroChamadoStr)
        if (len(temChamado) == 0):
            data_e_hora_atuais = datetime.now()
            DataCriacao = data_e_hora_atuais.strftime('%Y%m%d %H:%M:%S')
            Salva_registro = Grava_dados_ListaDistribuicao(numeroChamadoStr, opcao, opcaoNome1, opcaoNome2, opcaoNome3, membros, proprietarios, remetentes, aberta, fechada, StringEmailMembros, StringEmailProprietarios, DataCriacao, txtNomeCompletoMembros, txtNomeCompletoProprietarios, emailDaLista)
        else:
            Salva_registro = "OK"

        if 'ERRO' in Salva_registro:
            print('Houve ERRO na rotina de Grava_dados_ListaDistribuicao', Salva_registro)
            browser.close()
            return('Houve ERRO na rotina de Grava_dados_ListaDistribuicao',Salva_registro)
            sys.exit()
        else:
            return 'OK'
        browser.close()
    except Exception as e:
        PrintException()
        print('****ERRO NO PROCESSO DA ROTINA PEGA_DEMANDA****:', e) 
        return e