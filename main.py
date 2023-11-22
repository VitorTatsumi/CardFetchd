
############################################################################################################################
#### Definir como padrão (Excel):                                                                                       ####
#### NomeDoPokemon-VSTAR ao invés de NomeDoPokemon-V-Astro                                                              ####
#### Nomes em inglês. Ex: Flying Pikachu ao invés de Pikachu Voador                                                     ####    
#### Não faz busca por raridades acima de UR, nem por cartas PROMO                                                      ####
#### Não faz busca por nenhum tipo de carta além de Pokemons                                                            ####
####                                                                                                                    ####    
############################################################################################################################


############################################################################################################################
####                                                                                                                    ####    
####                                                                                                                    ####
####                                                                                                                    ####
####                                                                                                                    #### 
####                                 Desenvolvido por Vitor Tatsumi - 21/11/2023                                        ####
####                                                                                                                    ####
####                                                                                                                    ####    
####                                                                                                                    ####
####                                                                                                                    ####
####                                                                                                                    ####    
############################################################################################################################


import pyautogui
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

#Lendo a tabela
tabela = pd.read_excel('BotTesteLiga\pteste.xlsx')
df = pd.DataFrame(tabela, columns=['Nome','Código'])

#Exibindo o conteúdo das cédulas 1 das colunas parametrizadas
print(tabela['Nome'].values[1] + ' ' + tabela['Código'].values[1])
#var = 'https://www.ligapokemon.com.br/?view=cards/card&card=' + tabela['Nome'].values[1]

#print('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[1]) + '%20(' + str(tabela['Código'].values[1][1:4]) + '/'+ str(tabela['Código'].values[1][5:7]) + ')&ed=' + str(tabela['Coleção'].values[1]) + '&num=' + str(tabela['Código'].values[1][1:4]))

#Quantidade de cartas 
n = int(input('Insira a quantidade de cartas para consultar: '))

#Abre o navegador Chrome atribuindo-o à variável
navegador = webdriver.Chrome()


def searchCard():
    for i in range(n):
        #if (tabela['Nome'].values[1][-2:] == 'ex'):
        navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[i]) + '%20(' + str(tabela['Código'].values[i][1:4]) + '/'+ str(tabela['Código'].values[i][5:7]) + ')&ed=' + str(tabela['Coleção'].values[i]) + '&num=' + str(tabela['Código'].values[i]))
        #else:
            #navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[1]) + '%20(' + str(tabela['Código'].values[1][1:4]) + '/'+ str(tabela['Código'].values[1][5:7]) + ')&ed=' + str(tabela['Coleção'].values[1]) + '&num=' + str(tabela['Código'].values[1]))
            #navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[2]) + '%20(' + str(tabela['Código'].values[2][1:4]) + '/'+ str(tabela['Código'].values[2][5:7]) + ')&ed=' + str(tabela['Coleção'].values[2]) + '&num=' + str(tabela['Código'].values[2]))
        minVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[2]').get_attribute("textContent")
        midVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[4]').get_attribute("textContent")
        maxVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[6]').get_attribute("textContent")
        print('Menor valor: ' + minVal + '\nValor médio: ' + midVal + '\nMaior Valor: ' + maxVal)
        tabela.loc[tabela['Nome'] == str(tabela['Nome'].values[i]), 'Menor'] = str(minVal)
        tabela.loc[tabela['Nome'] == tabela['Nome'].values[i], 'Medio'] = str(midVal)
        tabela.loc[tabela['Nome'] == tabela['Nome'].values[i], 'Maximo'] = str(maxVal)

searchCard()
time.sleep(5)

os.remove('BotTesteLiga/ptesteNovo.xlsx')
tabela.to_excel('BotTesteLiga/ptesteNovo.xlsx', index=False)
time.sleep(100000)




############################################################################################################################
####                                                                                                                    ####    
####                                                                                                                    ####
####                                                                                                                    ####
####                                                                                                                    #### 
####                                 Desenvolvido por Vitor Tatsumi - 21/11/2023                                        ####
####                                                                                                                    ####
####                                                                                                                    ####    
####                                                                                                                    ####
####                                                                                                                    ####
####                                                                                                                    ####    
############################################################################################################################
