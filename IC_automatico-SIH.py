#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pyautogui
import time
import pyperclip
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

navegador = webdriver.Chrome(ChromeDriverManager().install())

plan = pd.read_excel(r"S:\Fechamentos\Fechamento_2022\MFG\Carga Componentes\IC Automatico.xlsm", sheet_name="Apoio ao IC")

num_ordem = int((plan['Apoio IC'] [0]))
Local = str((plan['Local'] [0]))

print(num_ordem)

# Abrir Chrome no JDE
#abre o link
navegador.get(r"https://e1lapd.aws.baxter.com/jde/E1Menu.maf?jdeowpBackButtonProtect=PROTECTED")
#Aguarda a pagina aparecer
while len(navegador.find_elements_by_id('carousel')) < 1:
    time.sleep(1)

    
wait = WebDriverWait(navegador, 20)

#Menu
wait.until(EC.element_to_be_clickable((By.ID, 'drop_mainmenu'))).click()

#baxter Menu
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div[2]/div[3]/div/div[2]/div[4]/div/div/div[2]/div[1]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span'))).click()

#Manufatura
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[10]/table/tbody/tr/td/div/div[1]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span'))).click()

#MFG Brasil
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[11]/table/tbody/tr/td/div/div/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span'))).click()

#Relatório chão de fabrica
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[12]/table/tbody/tr/td/div/div[3]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/span'))).click()

#Conclusões integrais
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[13]/table/tbody/tr/td/div/div[5]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/a'))).click()

#indo para o iFrame da pagina interna
navegador.switch_to.frame("e1menuAppIframe")

for x in range (num_ordem):
    wo = int((plan['Nº da Ordem'] [x]))
    time.sleep(1)
    
    
    #Tab do campo acima para o campo de ordem
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[3]/nobr/input'))).send_keys(Keys.TAB)
    
    #Campo WO
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table/tbody/tr[1]/td[2]/div/nobr/input'))).send_keys(wo)
    time.sleep(0.3)
    
    #Pesquisar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/a/img'))).click()
    time.sleep(1)
    
    #Seleciona todas
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table/tbody/tr[2]/td[1]/div[1]/input'))).click()
    time.sleep(0.3)
    
    #Selecionar(ENTRAR NA ORDEM)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img'))).click()
    
    Quant = str((plan['Quant. Pedida'] [x]))
    
    #Quantidade
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span/div/span[2]/table[2]/tbody/tr/td/div/span[6]/nobr/input'))).send_keys(Quant)

    #Aba Lote/Local
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/span/table/tbody/tr/td/span/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr/td/span/a'))).click()
    time.sleep(0.2)
    #Local
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span/div/span[2]/nobr/input'))).send_keys(Local)
    time.sleep(0.2)
    Lote = str((plan['Lote/ Série'] [x]))
    
    #TAB para campo lote
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span/div/span[2]/nobr/input'))).send_keys(Keys.TAB)
    time.sleep(0.2)
    #Lote
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/span/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span/div/span[6]/nobr/input'))).send_keys(Lote)
    
    #Salvar e Fechar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[3]/button'))).click()
    time.sleep(0.3)
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[3]/button'))).click()
    time.sleep(1)
    
    for b in range(8):
        #campo wo
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table/tbody/tr[1]/td[2]/div/nobr/input'))).send_keys(Keys.BACKSPACE)
    
    
    #Tab do campo acima para o campo de ordem
    #wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[1]/table[2]/tbody/tr/td/div/span[3]/nobr/input'))).send_keys(Keys.TAB)
    #time.sleep(0.3)

navegador.close()

# wait.until(EC.element_to_be_clickable((By.XPATH, ''))).click()
# wait.until(EC.element_to_be_clickable((By.XPATH, ''))).click()
# wait.until(EC.element_to_be_clickable((By.XPATH, ''))).click()
# wait.until(EC.element_to_be_clickable((By.XPATH, ''))).click()

#pyautogui.PAUSE = 1.0

#time.sleep(2)
#navegador.switch_to.default_content()

#By: Neyton Hugo


# In[ ]:





# In[ ]:





# In[ ]:




