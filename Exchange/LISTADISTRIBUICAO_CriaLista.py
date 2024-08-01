import pyodbc
import time, sys
import linecache
from selenium import webdriver #import webdriver module
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#LoginSharePoint
#userLogin = "automacao@m365labsuporte.onmicrosoft.com" #Conta
#userPass = "Fub92902" #Senha

userLogin = "SAMSAZU@petrobras.com.br" #Conta
userPass = "Automacao01@" #Senha

def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print ('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))

def MontaStringComEmailsO(EmailsSemDom, browser): #OWNER
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        #indice = EmailSemD.find('@')
        #EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Owners”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[4]/div/div/div/div[1]/div/div[1]/a').click()
        time.sleep(1)

        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[2]').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click()
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def MontaStringComEmailsM(EmailsSemDom, browser): #MEMBER
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        #indice = EmailSemD.find('@')
        #EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Members”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[5]/div[3]/div/div/div[1]/div/div[1]/a').click()
        time.sleep(1)
        
        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/a')
        userEmail.click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[2]/a').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[2]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click() #Salvar
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def MontaStringComEmailsO2(EmailsSemDom, browser): #OWNER 2
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        #indice = EmailSemD.find('@')
        #EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Owners”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/span[2]/a').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[2]/div/div[1]/div/div/div/div[1]/div/div[1]/a/img').click()
        time.sleep(1)

        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[2]').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click()
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def MontaStringComEmailsM2(EmailsSemDom, browser): #MEMBER 2
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        #indice = EmailSemD.find('@')
        #EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Members”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/span[3]/a').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[3]/div/div/div/div/div/div[1]/div/div[1]/a/img').click()
        time.sleep(1)
        
        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/a/img')
        userEmail.click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[2]/a').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[2]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click() #Salvar
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def ExcluiOwners(browser,Nomes):
    # Clica para abrir a janela de Proprietários
    source = browser.find_element(By.XPATH,'//*[@id="bookmarklink_1"]').click()
    number = 0 

    while number < len(Nomes) :
        try:
            time.sleep(1)
            WNome = Nomes[number]
            # Procura o nome para excluir
            source1 = browser.find_element(By.XPATH,'//td[@title="' + WNome + '"]').click()
            # Clica no icone de Excluir
            source1 = browser.find_element(By.XPATH,'//*[@id="ResultPanePlaceHolder_EditMailGroup_ownershipSection_contentContainer_ceOwner_listview_ToolBar"]/div[3]/a/img').click()

            number = number + 1
        except Exception as e:
            number = number + 1
            
    return 'OK'
    #return

def ExcluiMembers(browser,Nomes):
    # Clica para abrir a janela de Membros
    source = browser.find_element(By.XPATH,'//*[@id="bookmarklink_2"]').click()
    number = 0 
    
    while number < len(Nomes) :
        try:
            time.sleep(1)
            WNome = Nomes[number]
            # Procura o nome para excluir
            source1 = browser.find_element(By.XPATH,'//td[@title="' + WNome + '"]').click()
            # Clica no icone de Excluir
            source1 = browser.find_element(By.XPATH,'//*[@id="ResultPanePlaceHolder_EditMailGroup_MembershipSection_contentContainer_ceMembers_listview_ToolBar"]/div[3]/a/img').click()

            number = number + 1
        except Exception as e:
            number = number + 1
            
    return 'OK'
    #return
    
def Get_ListaDistribuicao(): 
    
    try:
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        cur = con.cursor()
        
        cur.execute("Select * from ListaDistribuicao where StatusCriaLista = 'PENDENTE'")
        
        dbData = cur.fetchall()
        #lista = dbData[0][0]
         
        cur.close()
        con.close()

        return dbData

    except Exception as e:
        print('****ERRO****:', e)

def ListaDistribuicaoInclusao(DadosBanco):
    
    try:
        browser = profile = webdriver.FirefoxProfile()
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))
        browser.get("https://outlook.office365.com/ecp")
    
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
        
        time.sleep(3)
        userEmail = browser.find_element(By.XPATH,'//*[@id="i0116"]')
        userEmail.send_keys(userLogin)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()

        time.sleep(3)
        userPsw = browser.find_element(By.XPATH,'//*[@id="i0118"]')
        userPsw.send_keys(userPass)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[10]/div[1]/ul/li[2]/span/a')))
        
        #Abrindo Grupo
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[1]/ul/li[2]/span/a').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[3]/div[2]/a').click()
        time.sleep(2)
        browser.switch_to.frame(browser.find_element(By.XPATH,'/html/body/form/div[10]/div[23]/iframe[1]'))
        
        #TAB + ENTER na adição
        time.sleep(15)
        browser.find_element(By.XPATH,'//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar_NewDistributionGroupSplitButton"]/table/tbody/tr/td/div/a/img').click()
        time.sleep(2)
        action = ActionChains(browser)
        action.send_keys(Keys.TAB)
        action.send_keys(Keys.ENTER)
        action.perform()
        
        time.sleep(3)
        # Muda para a nova Janela
        window_before = browser.window_handles[0]
        window_after = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela
        
        time.sleep(2)
        arrayName = Get_ListaDistribuicao()
        opcaoName = arrayName[0][3]
        opcaoName = opcaoName.strip()
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
        userName.click()
        userName.send_keys(opcaoName) #campo DisplayName
        
        time.sleep(2)            
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
        userName.click()
        userName.send_keys(opcaoName) #campo Alias
        
        time.sleep(2)
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input').click() #Campo Email

        # Ativar em produção
        time.sleep(1)
        userName = browser.find_element(By.XPATH,'//*[@id="ResultPanePlaceHolder_Wizard1_generalInfoSection_contentContainer_lgvAlias_listDomain"]/option[2]').click() #Campo Dominio

        #Procurando email sem @dominio em Owner
        time.sleep(1)
        #campo2 = arrayName[0][9] #Pegando info do campo proprietarios
        campo2 = str(arrayName[0][17]) #Pegando as CHAVES dos Proprietários (são separadas por virgula)
        campo2 = campo2.rstrip()
        ownerName = campo2.split('|')
        MontaStringComEmailsO(ownerName, browser)
        
        #Procurando email sem @dominio em Membros
        time.sleep(1)
        #campo1 = arrayName[0][7] #Pegando info do campo membros
        campo1 = str(arrayName[0][16]) #Pegando as CHAVES dos Membros (são separadas por virgula)
        campo1 = campo1.rstrip()
        memberName = campo1.split('|')
        MontaStringComEmailsM(memberName, browser)
               
        #Opção entrada Grupo
        DadosBanco = Get_ListaDistribuicao()       
        opnBtn = DadosBanco
        opcao = str(opnBtn[0][10])       
        opcao = opcao.strip()
              
        if (opcao == 'Aberta'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[1]/table/tbody/tr[1]/td/input').click()

        if (opcao == 'Fechada'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[1]/td/input').click()
        
        if (opcao == 'Owner Approval'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[1]/table/tbody/tr[3]/td/input').click()   
            
        #Opção saída Grupo
        opcao2 = str(opnBtn[0][11])
        opcao2 = opcao2.strip()
        
        if (opcao2 == 'Aberta'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[1]/td/input').click()
            
        if (opcao2 == 'Fechada'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[2]/td/input').click()      

        salvaEmail = opcaoName 

        #Salvando
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
        time.sleep(5)
        
        ############################################ LOOPING PARA CASO DE ERRO - EMAIL CADASTRADO 2 ############################################
        
        #Mensagem de erro
        txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[1]/div/div/div/span[1]')#Texto do Erro
        texto = txtMensage.text
        
        #txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button').click() - xpath botão OK
        #time.sleep(2)    
    
        if ('There are multiple recipients matching identity' in texto or 'Há vários destinatários que correspondem à identidade' in texto):
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button')
            txtName.click()
            time.sleep(1)
            
            arrayName = Get_ListaDistribuicao()
            opcaoName2 = arrayName[0][4]
            opcaoName2 = opcaoName2.strip()
            
            #Campo Display Name
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
            txtName.click()
            txtName.clear()
            txtName.send_keys(opcaoName2)
            time.sleep(1)

            #Campo Alias
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName2)
            time.sleep(1)
            
            #Campo Email
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName2)
            time.sleep(1)
            
            salvaEmail = opcaoName2 

            #Salvando
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
            time.sleep(5)
            
        ############################################ LOOPING PARA CASO DE ERRO - EMAIL CADASTRADO 3 ############################################
        
        #Mensagem de erro
        txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[1]/div/div/div/span[1]')#Texto do Erro
        texto = txtMensage.text
        
        #txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button').click() - xpath botão OK
        #time.sleep(2)    
    
        if ('There are multiple recipients matching identity' in texto or 'Há vários destinatários que correspondem à identidade' in texto):
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button')
            txtName.click()
            time.sleep(1)
            
            arrayName = Get_ListaDistribuicao()
            opcaoName3 = arrayName[0][5]
            opcaoName3 = opcaoName3.strip()
            
            #Campo Display Name
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
            txtName.click()
            txtName.clear()
            txtName.send_keys(opcaoName3)
            time.sleep(1)

            #Campo Alias
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName3)
            time.sleep(1)
            
            #Campo Email
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName3)
            time.sleep(1)
            
            salvaEmail = opcaoName3

            #Salvando
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
            time.sleep(5)
            
        browser.close()
        browser.switch_to.window(window_before)
        browser.close()
        
        return 'OK', salvaEmail
    except Exception as e:
        PrintException()

        print('****ERRO NO FILTRO****:', e)
        return '****ERRO NO FILTRO****:', e

def ListaDistribuicaoIncluiChaves(DadosBanco):
    try:
        browser = profile = webdriver.FirefoxProfile()
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))
        browser.get("https://outlook.office365.com/ecp")
    
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
        
        time.sleep(3)
        userEmail = browser.find_element(By.XPATH,'//*[@id="i0116"]')
        userEmail.send_keys(userLogin)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()

        time.sleep(3)
        userPsw = browser.find_element(By.XPATH,'//*[@id="i0118"]')
        userPsw.send_keys(userPass)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[10]/div[1]/ul/li[2]/span/a')))
        
        #Abrindo Grupo
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[1]/ul/li[2]/span/a').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[3]/div[2]/a').click()
        time.sleep(10)
        browser.switch_to.frame(browser.find_element(By.XPATH,'/html/body/form/div[10]/div[23]/iframe[1]'))
                
        #Buscar Email do Grupo para Alterar
        arrayName = Get_ListaDistribuicao()
        opcaoName1 = str(arrayName[0][15])
        opcaoName1 = opcaoName1.strip()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/a/img').click()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[1]/input')
        source.click()
        source.send_keys(opcaoName1)
        
        #Double Click no grupo
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[2]/a/img').click()
        time.sleep(2)
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[2]/div[2]/div[3]/table/tbody/tr/td[1]')     
        action = ActionChains(browser)
        action.double_click(source)
        action.perform()
        
        # Muda para a nova Janela
        window_before = browser.window_handles[0]
        window_after = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela
        
        #Procurando email sem @dominio em Owner
        time.sleep(1)
        #campo2 = arrayName[0][9] #Pegando info do campo proprietarios
        campo2 = str(arrayName[0][17]) #Pegando as CHAVES dos Proprietários (são separadas por virgula)
        if (campo2 != 'None'): 
            campo2 = campo2.rstrip()
            ownerName = campo2.split('|')
            MontaStringComEmailsO2(ownerName, browser)
        
        #Procurando email sem @dominio em Membros
        time.sleep(1)
        #campo1 = arrayName[0][7] #Pegando info do campo membros
        campo1 = str(arrayName[0][16]) #Pegando as CHAVES dos Membros (são separadas por virgula)
        campo1 = campo1.rstrip()
        memberName = campo1.split('|')
        MontaStringComEmailsM2(memberName, browser)
               
        #Opção entrada Grupo
        DadosBanco = Get_ListaDistribuicao()       
        opnBtn = DadosBanco
        opcao = str(opnBtn[0][10])       
        opcao = opcao.strip()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/span[4]/a').click()
              
        if (opcao == 'Aberta'):
            source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[4]/div/div/div[1]/table/tbody/tr[1]/td/input').click()

        if (opcao == 'Fechada'):
            source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[4]/div/div/div[1]/table/tbody/tr[2]/td/input').click()
        
        if (opcao == 'Owner Approval'):
            source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[4]/div/div/div[1]/table/tbody/tr[3]/td/input').click()   
            
        #Opção saída Grupo
        opcao2 = str(opnBtn[0][11])
        opcao2 = opcao2.strip()
        
        if (opcao2 == 'Aberta'):
            source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[4]/div/div/div[2]/table/tbody/tr[1]/td/input').click()
            
        if (opcao2 == 'Fechada'):
            source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/div/div/div[4]/div/div/div[2]/table/tbody/tr[2]/td/input').click()      
            
        #Salvando
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[4]/button[1]').click()
        time.sleep(5)
            
        Whandles = browser.window_handles
        browser.close()
        browser.switch_to.window(window_before)
        browser.close()
        
        return 'OK'
    except Exception as e:
        print('****ERRO NA ALTERAÇÃO ****:', e)
        return '****ERRO NA ALTERAÇÃO ****:', e

def ListaDistribuicaoExcluir(DadosBanco):
    try:
        browser = profile = webdriver.FirefoxProfile()
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))
        browser.get("https://outlook.office365.com/ecp")
    
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
        
        #LoginSharePoint
        userLogin = "SAMSAZU@petrobras.com.br" #Conta
        userPass = "Automacao01@" #Senha
        
        time.sleep(3)
        userEmail = browser.find_element(By.XPATH,'//*[@id="i0116"]')
        userEmail.send_keys(userLogin)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()

        time.sleep(3)
        userPsw = browser.find_element(By.XPATH,'//*[@id="i0118"]')
        userPsw.send_keys(userPass)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[10]/div[1]/ul/li[2]/span/a')))
        
        #Abrindo Grupo
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[1]/ul/li[2]/span/a').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[3]/div[2]/a').click()
        time.sleep(10)
        browser.switch_to.frame(browser.find_element(By.XPATH,'/html/body/form/div[10]/div[23]/iframe[1]'))
        
        #Buscar Email do Grupo para Alterar
        arrayName = Get_ListaDistribuicao()
        opcaoName1 = str(arrayName[0][15])
        opcaoName1 = opcaoName1.strip()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/a/img').click()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[1]/input')
        source.click()
        source.send_keys(opcaoName1)
        
        #Double Click no grupo
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[2]/a/img').click()
        time.sleep(2)
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[2]/div[2]/div[3]/table/tbody/tr/td[1]')     
        action = ActionChains(browser)
        action.double_click(source)
        action.perform()
        
        # Muda para a nova Janela
        window_before = browser.window_handles[0]
        window_after = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela
        time.sleep(2)

        #Excluindo Nomes do campo Owner
        WOwner = str(arrayName[0][17])
        ownerName = WOwner.split('|')
        if (len(ownerName) > 0):
            ExcluiOwners(browser, ownerName)
                        
        #Excluindo Dados do campo Membros
        time.sleep(1)
        WMembers = str(arrayName[0][16])
        WMembers = WMembers.split('|')
        if (len(WMembers) > 0):
            ExcluiMembers(browser, WMembers)
        
        #Salvando
        browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[4]/button[1]').click()
        time.sleep(2)
        browser.close()
        browser.switch_to.window(window_before)
        browser.close()
        
        return 'OK'
    except Exception as e:
        print('****ERRO NA EXCLUSÃO DA CHAVE****:', e)
        return '****ERRO NA EXCLUSÃO DA CHAVE****', e

def AtualizaStatus(numeroChamado, EmailUsado):
    try:
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        cur = con.cursor()
        
        if (EmailUsado == ''):
            cur.execute("Update ListaDistribuicao Set StatusCriaLista = 'OK' where NumeroChamado = " + "'" + numeroChamado + "'")
        else:
            cur.execute("Update ListaDistribuicao Set StatusCriaLista = 'OK', EmailDaLista = " + "'" + EmailUsado + "'" +  "where NumeroChamado = " + "'" + numeroChamado + "'")

        con.commit()          
        cur.close()
        con.close()

        return 'OK'
    except Exception as e:
        print('****ERRO****:', e) 

def ListaDistribuicaoExcluirLista(DadosBanco):
    try:
        browser = profile = webdriver.FirefoxProfile()
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))
        browser.get("https://outlook.office365.com/ecp")
    
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
        
        #LoginSharePoint
        userLogin = "SAMSAZU@petrobras.com.br" #Conta
        userPass = "Automacao01@" #Senha
        
        time.sleep(3)
        userEmail = browser.find_element(By.XPATH,'//*[@id="i0116"]')
        userEmail.send_keys(userLogin)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()

        time.sleep(3)
        userPsw = browser.find_element(By.XPATH,'//*[@id="i0118"]')
        userPsw.send_keys(userPass)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[10]/div[1]/ul/li[2]/span/a')))
        
        #Abrindo Grupo
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[1]/ul/li[2]/span/a').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[3]/div[2]/a').click()
        time.sleep(15)
        browser.switch_to.frame(browser.find_element(By.XPATH,'/html/body/form/div[10]/div[23]/iframe[1]'))
        
        #Buscar grupo para Exclusão
        arrayName = Get_ListaDistribuicao()
        opcaoName1 = str(arrayName[0][15])
        opcaoName1 = opcaoName1.strip()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/a/img').click()
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[1]/input')
        source.click()
        source.send_keys(opcaoName1)
        
        #Exclusão do Grupo
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[19]/div/table/tbody/tr/td[2]/a').click()
        time.sleep(1)
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[2]/div[2]/div[3]/table/tbody/tr[1]/td[1]').click()
        time.sleep(1)
        source = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div/div[2]/table/tbody/tr/td[1]/div/div[1]/div/div[11]/a').click()
        time.sleep(2)

        # Clica no SIM na janela de confirmação
        browser.switch_to.default_content()
        browser.find_element_by_class_name('btnbox').click()
        time.sleep(5)
            
        browser.close()
        
        return 'OK'
    except Exception as e:
        print('****ERRO NA EXCLUSÃO DA LISTA****:', e)
        return '****ERRO NA EXCLUSÃO DA LISTA****', e

def RotinaPrincipal():
    DadosBanco = Get_ListaDistribuicao()
    EmailUsado = ''
    RetornoInclusao = ''
        
    if (len(DadosBanco) > 0 ):
        if DadosBanco[0][2] == "Criar (menos que 100 membros)":
           RetornoInclusao, EmailUsado = ListaDistribuicaoInclusao(DadosBanco)
           if RetornoInclusao == 'OK':
              AtualizaStatus(DadosBanco[0][1], EmailUsado)
              return 'OK'
            
        if DadosBanco[0][2] == "Incluir chaves":
           RetornoAlteracao = ListaDistribuicaoIncluiChaves(DadosBanco)
           if RetornoAlteracao == 'OK':
              AtualizaStatus(DadosBanco[0][1],'')
              return 'OK'
        
        if DadosBanco[0][2] == "Excluir chaves":
           RetornoExclusao = ListaDistribuicaoExcluir(DadosBanco)
           if RetornoExclusao == 'OK':
              AtualizaStatus(DadosBanco[0][1],'')
              return 'OK'

        if DadosBanco[0][2] == "Excluir lista":
           RetornoExclusaoLista = ListaDistribuicaoExcluirLista(DadosBanco)
           if RetornoExclusaoLista == 'OK':
              AtualizaStatus(DadosBanco[0][1],'')
              return 'OK'
        
        # --------------------------   Identificar OPÇÃO INVALIDA ------------
        
        #if DadosBanco[0][2] == "Excluir"    
        #ListaDistribuicaoExcluir(DadosBanco)

#SRotinaPrincipal()