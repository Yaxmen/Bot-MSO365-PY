import time
from selenium import webdriver
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
        
def Filtrar(Filtro,browser):
    try:
        time.sleep(10)
        element = WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="queryProfileSummaryCell_4"]/div[1]/a')))
        browser.find_element_by_xpath('//*[@id="queryProfileSummaryCell_4"]/div[1]/a').click()
        #queryProfileList
        
        time.sleep(10)
        # Aguarda o botão de AÇÕES aparecer
        WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="queryProfileList"]')))
        #filtro1 = WebDriverWait(browser,30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="queryProfileList"]')))
        browser.find_element_by_xpath('//*[@id="queryProfileList"]').click()
        
        time.sleep(10)
        # Aguarda a tela de filtros aparecer
        WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="filterCriteria_tablist_classificationTab"]')))
        browser.find_element_by_xpath('//*[@id="filterCriteria_tablist_classificationTab"]').click()
        
        time.sleep(10)
        # Digita o filtro solicitado
        userName = browser.find_element_by_xpath('//*[@id="EventSearchForm_NONE_event_lookup_serviceOffering_textNode"]')
        userName.click()
        userName.send_keys(Filtro)
        time.sleep(1)
        userName = browser.find_element_by_xpath('//*[@id="btExecute"]/span[1]').click()
        time.sleep(3)
        
        # CLica no botão Filtrar 
        userName = browser.find_element_by_xpath('//*[@id="btExecute"]/span[1]').click()
        
        return 'OK'
    except Exception as e:
        print('****ERRO NA ROTINA DE FILTRAR****:', e) 
        return e
