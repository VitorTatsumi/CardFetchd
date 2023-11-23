
############################################################################################################################
#### Definir como padrão (Excel):                                                                                       ####
#### NomeDoPokemon-VSTAR ao invés de NomeDoPokemon-V-Astro                                                              ####
#### Nomes em inglês. Ex: Flying Pikachu ao invés de Pikachu Voador                                                     ####    
#### Não faz busca por raridades acima de UR, nem por cartas PROMO                                                      ####
#### Não faz busca por nenhum tipo de carta além de Pokemons                                                            ####
#### Pokemons EX antigos deverão ser inseridos com -EX e Pokemons ex novos deverão ser inseridos sem o    ############################################################################################################################


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
from tkinter import *
from tkinter import ttk

##################################### Início de c
#Criação da janela
tela = Tk()
#Edição do título
tela.title('Robô Liga Pókemon')

texto = Label(tela, text="Clique abaixo para iniciar o robô")
texto.grid(column=0, row= 0)

frm = ttk.Frame(tela, padding=10)


frm.grid()
ttk.Label(frm, text="Hello World!").grid(column=0, row=0)
ttk.Button(frm, text="Quit", command=tela.destroy).grid(column=1, row=0)







#Lendo o arquivo Excel 
tabela = pd.read_excel('BotTesteLiga\pteste.xlsx')
df = pd.DataFrame(tabela, columns=['Nome','Código', 'Coleção', 'Menor', 'Medio', 'Maximo'])

#Quantidade de cartas 
n = int(input('Qual a quantidade de cartas para consulta?\n '))

#Abre o navegador Chrome atribuindo-o à variável
navegador = webdriver.Chrome()

#Função para busca e armazenamento das cartas na tabela
def searchCard(qtd):
    #Laço de repetição para verificação da quantidade de cartas previamente inserida
    for i in range(qtd):

        #O navegador faz a busca diretamente pela URL, com os parâmetros das cartas (Nome, código, coleção, código)
        navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[i]) + '%20' + str(tabela['Código'].values[i]) + '&ed=' + str(tabela['Coleção'].values[i]) + '&num=' + str(tabela['Código'].values[i]))
        #TryCatch para que o programa não seja fechado ao encontrar um erro, mas sim que exiba uma mensagem informando
        try:
            #Através do elemento XPATH, o navegador encontra o objeto de texto e seu conteúdo. Neste caso, os valores das cartas, mínimo, médio e máximo.
            minVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[2]').get_attribute("textContent")
            midVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[4]').get_attribute("textContent")
            maxVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[6]').get_attribute("textContent")
            #Pandas localiza a célula correta à se fazer atribuição do valor das cartas. Neste caso, pegando os valores de nome, código e coleção da carta atual e comparando ela com os valores dentro das variáveis 
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]) , 'Menor']  = str(minVal)
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Medio'] = str(midVal)
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Maximo'] = str(maxVal)
        except:
            print ('Pókemon: ' + str(tabela['Nome'].values[i] + str(tabela['Código'].values[i] + ' Não foi lido.')))
    return time.sleep(1)


#Chama a função que busca e armazena a carta na tabela
searchCard(n)
navegador.quit()
print('Busca realizada.\nDados salvos em ptesteNovo.xlsx')
os.remove('BotTesteLiga/ptesteNovo.xlsx')
tabela.to_excel('BotTesteLiga/ptesteNovo.xlsx', index=False)

#Deixa a tela exibida
tela.mainloop()




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
