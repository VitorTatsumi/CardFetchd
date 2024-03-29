
############################################################################################################################
#### Definir como padrão (Excel):                                                                                       ####
#### NomeDoPokemon-VSTAR ao invés de NomeDoPokemon-V-Astro                                                              ####
#### Nomes em inglês. Ex: Flying Pikachu ao invés de Pikachu Voador                                                     ####    
####                                                                                                                    ####
####                                                                                                                    ####
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
from tkinter import *
from tkinter import ttk



#Lendo o arquivo Excel 
tabela = pd.read_excel('CardFetchd\CardFetchd_Principal.xlsx')
df = pd.DataFrame(tabela, columns=['Nome','Código', 'Coleção', 'Menor', 'Medio', 'Maximo', 'Qualidade'])

#Quantidade de cartas, verifica pela quantidade de linhas da tabela
#n = int(input('Qual a quantidade de cartas para consulta?\n '))
n = len(df.index)



#Função para busca e armazenamento das cartas na tabela
def searchCard():
    #Limpa o histórico de erros
    pokeError["text"] == ''

    #Abre o navegador Chrome atribuindo-o à variável
    navegador = webdriver.Chrome()
    #Laço de repetição para verificação da quantidade de cartas previamente inserida
    for i in range(n):
        #O navegador faz a busca diretamente pela URL, com os parâmetros das cartas (Nome, código, coleção, código)
        navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[i]) + '%20' + str(tabela['Código'].values[i]) + '&ed=' + str(tabela['Coleção'].values[i]) + '&num=' + str(tabela['Código'].values[i]))
        #TryCatch para que o programa não seja fechado ao encontrar um erro, mas sim que exiba uma mensagem informando
        try:
            #Através do elemento XPATH, o navegador encontra o objeto de texto e seu conteúdo. Neste caso, os valores das cartas, mínimo, médio e máximo.
            minVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[2]').get_attribute("textContent")
            midVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[4]').get_attribute("textContent")
            maxVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[6]').get_attribute("textContent")
            try:
                qualidade = navegador.find_element('xpath', '/html/body/main/div[6]/div[3]/div[1]/div[5]/div[5]').get_attribute("textContent")
            except:
                try:
                    qualidade = navegador.find_element('xpath', '/html/body/main/div[6]/div[3]/div[1]/div[6]/div[5]').get_attribute("textContent")
                except:
                    try:
                        qualidade = navegador.find_element('xpath', '/html/body/main/div[7]/div[3]/div[1]/div[5]/div[5]').get_atribute("textContent")
                    except:
                        try:
                            qualidade = navegador.find_element('xpath', '/html/body/main/div[7]/div[3]/div[1]/div[6]/div[5]').get_atribute("textContent")
                        except:
                            try:
                                qualidade = navegador.find_element('xpath', '/html/body/main/div[7]/div[3]/div[1]/div[6]/div[5]').get_atribute("textContent")
                            except: 
                                qualidade = 'N/E'

            #Pandas localiza a célula correta à se fazer atribuição do valor das cartas. Neste caso, pegando os valores de nome, código e coleção da carta atual e comparando ela com os valores dentro das variáveis  
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]) , 'Menor']  = str(minVal[3:])
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Medio'] = str(midVal[3:])
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Maximo'] = str(maxVal[3:])
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Qualidade'] = str(qualidade)

        except:
            pokeError["text"] += 'Pókemon: ' + str(tabela['Nome'].values[i] + str(tabela['Código'].values[i] + ' Não foi lido.\n'))
            #print ('Pókemon: ' + str(tabela['Nome'].values[i] + str(tabela['Código'].values[i] + ' Não foi lido.')))
    navegador.quit()
    #Define a mensagem de texto que aparecerá ao fim da busca
    mensagem["text"] = 'Busca realizada.\nDados salvos em CardFetchdResultsXLSX.xlsx\nDados salvos em CardFetchdResultsCSV.csv'
    os.remove('CardFetchd/CardFetchdResultsXLSX.xlsx')
    os.remove('CardFetchd/CardFetchdResultsCSV.csv')

    tabela.to_csv('CardFetchd/CardFetchdResultsCSV.csv', index=False)
    tabela.to_excel('CardFetchd/CardFetchdResultsXLSX.xlsx', index=False)
    return #time.sleep(1)

#Chama a função que busca e armazena a carta na tabela
#searchCard(n)

##################################### Início de c
#Criação da janela
#Interface gráfica TKinter
tela = Tk()
#Edição do título
tela.title("CardFetch'd")
#Tamanho inicial da tela
#tela.geometry("250x200")
#Elementos da tela
#Grid para determinar onde serão localizados os ite

texto = Label(tela, text="Clique abaixo para iniciar a busca de cartas").grid(column=0, row= 1, padx=10, pady=10)
botaoPesquisa = Button(tela, text="Pesquisar", command=searchCard).grid(column=0, row=2, padx=10, pady=10)
botaoFechar = Button(tela, text="Fechar", command=tela.destroy).grid(column=0, row=3, padx=10, pady=10)
pokeError = Label(tela, text="")
pokeError.grid(column=0, row=5, padx=10, pady=10)
mensagem = Label(tela, text="")
mensagem.grid(column=0, row=6, padx=10, pady=10)
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
