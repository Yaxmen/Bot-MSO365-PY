from time import time
from selenium import webdriver #import webdriver module
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
import time
import pyodbc
import pydirectinput
import keyboard
import linecache,sys

def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print ('EXCEPTION IN (LINE {} "{}"): {}'.format(lineno, line.strip(), exc_obj))

def GetCredentials(UserLogin):
    
    try:
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        cur = con.cursor()
        cur.execute("EXEC Get_Credentials '" + UserLogin + "'")
        
        dbData = cur.fetchall()
        loginPassword = dbData[0][0]
        cur.close()
        con.close()

        return loginPassword
    except Exception as e:
        print('****ERRO GETCREDENTIAL****:', e)

def AtualizaStatus(numeroChamado):
    try:
        con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        #con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        cur = con.cursor()
        
        cur.execute("Update ListaDistribuicao Set StatusFechaChamado = 'OK' where NumeroChamado = " + "'" + numeroChamado + "'")

        con.commit()          
        cur.close()
        con.close()

        return 'OK'
    except Exception as e:
        print('****ERRO****:', e) 

def FechaChamado(NumeroChamado, Departamento, Descricao, CausaItem, Causa):
    try:
        options = webdriver.ChromeOptions()
        preferences = {"profile.default_content_setting_values.notifications" : 2}
        options.add_experimental_option("prefs", preferences)

        browser = profile = webdriver.FirefoxProfile()
        #browser = webdriver.Chrome(executable_path=(r'C:\Program Files\WebDrivers\chromedriver.exe'),chrome_options=options)
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Program Files\WebDrivers\geckodriver.exe'))
        
        # Login
        userLogin = "SAMSAZU@petrobras.com.br"
        #userPass = GetCredentials(userLogin)
        userPass = 'Automacao01@'
        
        browser.get("http://gestaoclick.petrobras.com.br/")
        time.sleep(3)
        
        # Faz o LOGIN na tela do Navegador
        pydirectinput.moveTo(790, 200)
        pydirectinput.click()
        keyboard.write(userLogin)
        time.sleep(1)
        pydirectinput.press('tab')
        time.sleep(1)
        keyboard.write(userPass) 
        time.sleep(1)
        pydirectinput.press('tab')
        time.sleep(1)
        pydirectinput.press('return')
        
        time.sleep(20)
                
        # Aguarda aparecer a tela 
        element = WebDriverWait(browser,60).until(EC.presence_of_element_located((By.XPATH, '//*[@id="queryProfileSummaryCell_4"]')))
        browser.find_element(By.XPATH,'//*[@id="queryProfileSummaryCell_4"]').click()
        
        element = WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="queryProfileList"]')))

        # Pesquisa pelo número do Chamado
        numeroDoChamado = NumeroChamado
        time.sleep(3)
        CampoNumeroChamado = browser.find_element(By.XPATH,'//*[@id="components_TextBox_0"]')
        CampoNumeroChamado.send_keys(numeroDoChamado)
        # Clica no botão PESQUISAR
        time.sleep(2)
        btnNext = browser.find_element(By.XPATH,'//*[@id="dijit_form_ComboButton_0_button"]/div[1]').click()
       
        # Aguarda aparecer a segunda janela com os dados do Chamado
        time.sleep(10)
        Whandles = browser.window_handles
        
        window_before = browser.window_handles[0]
        window_after = browser.window_handles[1]
        browser.switch_to.window(window_after)

        element = WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="dojox_grid__View_1"]/div/div/div/div/table/tbody/tr/td[2]')))

        # Botao Direito no chamado para abrir opções 
        action = ActionChains(browser)
        action.move_to_element(browser.find_element(By.XPATH,'//*[@id="dojox_grid__View_1"]/div/div/div/div/table/tbody/tr/td[2]')).perform()
        action.context_click().perform() 

        # Clica na opção de "Exibir Detalhes do Evento"
        time.sleep(2)
        #browser.find_element(By.XPATH,'//*[@id="components_MenuItem_106_text"]').click()
        browser.find_element(By.XPATH,'//body[1]/div[24]/table[1]/tbody[1]/tr[6]/td[2]').click()
        

        time.sleep(5)
        window_Popup = browser.window_handles[2]
        browser.switch_to.window(window_Popup)

        Whandles = browser.window_handles

        # Pega o número da Tarefa da janela de Detalhes
        time.sleep(10)
        c0 = browser.find_element(By.XPATH,'//*[@id="customFieldsPlacmentNode"]/table[2]/tbody/tr[3]/td[3]')
        numeroTarefa = c0.text
        print('Tarefa ======> ' , numeroTarefa, '<=======')
        browser.close()

        # Volta para a Janela de pesquisa e fecha
        browser.switch_to.window(window_after)
        
        # Clica em "Assinalar para mim" (Chamado) 
        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@id="assignToMeAction"]/span[1]').click()
        time.sleep(30)

        # Fecha janela de pesquisa
        browser.close()

        time.sleep(2)
        browser.switch_to.window(window_before)
        # Pesquisa pelo número da Tarefa
        campoPesquisa = browser.find_element(By.XPATH,'//*[@id="components_TextBox_0"]')
        campoPesquisa.clear()
        campoPesquisa.send_keys(numeroTarefa)
        # Clica no botão PESQUISAR
        time.sleep(2)
        btnNext = browser.find_element(By.XPATH,'//*[@id="dijit_form_ComboButton_0_button"]/div[1]').click()

        time.sleep(30)
        Whandles1 = browser.window_handles
        window_after = browser.window_handles[1]
        
        time.sleep(5)
        browser.switch_to.window(window_after)
        browser.find_element(By.XPATH,'//*[@id="assignToMeAction"]/span[1]').click()
        time.sleep(30)

        # CLica no botão de AÇÕES
        browser.find_element(By.XPATH,'//*[@id="menuActions_label"]').click()

        # Seleciona a opção RESOLVER
        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@id="menuActions_$DisplayOnceAction(PENDING_CLOSURE)_text"]').click()

        # Muda o foco para a janela RESOLVER
        time.sleep(5)
        Whandles2 = browser.window_handles
        window_Resolver = browser.window_handles[1]
        browser.switch_to.window(window_Resolver)

        # Preenche o Departamento de Serviço Atribuido
        time.sleep(2)
        departamentoAtribuido = browser.find_element(By.XPATH,'//*[@id="ManageActionForm_NONE_assignedServDept_textNode"]')
        departamentoAtribuido.send_keys(Departamento)
        time.sleep(2)
        # Clica para confirmar a seleção
        confirma = browser.find_element(By.XPATH,'//*[@id="actionPane"]/table/tbody/tr/td/div[1]').click()
        
        # Preenche a Descrição
        time.sleep(2)
        ckeditor_frame = browser.find_element(By.CLASS_NAME, 'cke_wysiwyg_frame')
        browser.switch_to.frame(ckeditor_frame)
        editor_body = browser.find_element(By.TAG_NAME, 'body')
        editor_body.send_keys(Descricao)
        # Volta para o iFrame do formulário
        browser.switch_to.window(window_Resolver)
       
        # Preenche o Item de Causa
        time.sleep(2)
        causaItem = browser.find_element(By.XPATH,'//*[@id="ManageActionForm_NONE_causeItem_textNode"]')
        causaItem.send_keys(CausaItem)
        time.sleep(1)
        causaItemConfirma = browser.find_element(By.XPATH,'//*[@id="ManageActionForm_NONE_causeItem_comboBox_popup0"]').click()
        
        # Preenche o Causa
        time.sleep(2)
        campoCausa = browser.find_element(By.XPATH,'//*[@id="ManageActionForm_NONE_causeCategory_textNode"]')
        campoCausa.send_keys(Causa)
        time.sleep(2)
        campoCausaConfirma = browser.find_element(By.XPATH,'//*[@id="ManageActionForm_NONE_causeCategory_comboBox_popup0"]').click()
       
        # Clica no botão SALVAR AÇÃO
        time.sleep(2)
        browser.find_element(By.XPATH,'//*[@id="ManageActionForm.btSave_label"]').click()
        time.sleep(10)

        browser.close()
        browser.switch_to.window(window_before)
        browser.close()
        #print('Chamado ======> ' , numeroDoChamado, '<======= Fechado com Sucesso')
        
        AtualizaStatus(numeroDoChamado)
        #browser.close()
        return 'OK'        
    except Exception as e:
        print('****ERRO****:', e)
        #Whandles = browser.window_handles
        browser.close()
        #browser.switch_to.window(window_before)
        #browser.close()
        return ('ERRO na rotina de Fechar chamado')

#FechaChamado('S3149397', 'N3-MICROSOFT_OFFICE', 'A solicitação foi devidamente resolvida', 'Office 365', 'Outros')
