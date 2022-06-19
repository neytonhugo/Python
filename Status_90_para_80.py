#!/usr/bin/env python
# coding: utf-8

# In[6]:


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

plan = pd.read_excel(r"S:\Neyton Hugo\Voltar_ordens_80.xlsx")

print(plan['Apoio ao Python'] [0])

num_ordem = int((plan['Apoio ao Python'] [0]))

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

#Chance Work Status
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[13]/table/tbody/tr/td/div/div[8]/div/table/tbody/tr/td[4]/table/tbody/tr/td/table/tbody/tr/td[1]/a'))).click()

#indo para o iFrame da pagina interna
navegador.switch_to.frame("e1menuAppIframe")

for x in range (num_ordem):
    wo = int((plan['Ordens'] [x]))
    time.sleep(1)
    #Campo W.O
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table/tbody/tr[1]/td[2]/div/nobr/input'))).send_keys(wo)
    
    #pesquisar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[3]/a/img'))).click()
    
    #Select / Entrar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img'))).click()
    
    #Campo Status
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span/table[2]/tbody/tr/td/div/span[14]/nobr/input'))).send_keys("80")
    
    #Select / Fechar
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/table[2]/tbody/tr/td/table/tbody/tr/td[1]/a/img'))).click()
    
    #Retorna o campo ao estado de seleção
    wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[3]/div/table/tbody/tr/td/div/span[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[1]/td/div[1]/table/tbody/tr[1]/td[2]/div/nobr/input'))).send_keys(Keys.RETURN)
    





# In[3]:


#for y in range (num_ordem):
#    wos = int((plan['Ordens'] [y]))
#    pyperclip.copy(wos)
#    print(wos)





# In[8]:





# In[ ]:




