import pyautogui
import pandas as pd
import selenium 
import time
from selenium import webdriver


tabela = pd.read_excel('BotTesteLiga\pteste.xlsx')
print (tabela)


navegador = webdriver.Chrome()
navegador.get('https://www.ligapokemon.com.br/')
time.sleep(2)
click_button = navegador.find_element('xpath', '/html/body/div[3]/div[1]/img')#.send_keys('Cotação dólar'))
click_button.click()
busca = navegador.find_element('xpath', '/html/body/header/nav[1]/div/div[1]/div[2]/form/input[2]').send_keys('Cotação dólar')
time.sleep(100000)
