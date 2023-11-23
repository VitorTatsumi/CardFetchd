
############################################################################################################################
#### Definir como padrão (Excel):                                                                                       ####
#### NomeDoPokemon-VSTAR ao invés de NomeDoPokemon-V-Astro                                                              ####
#### Nomes em inglês. Ex: Flying Pikachu ao invés de Pikachu Voador                                                     ####    
#### Não faz busca por raridades acima de UR, nem por cartas PROMO                                                      ####
#### Não faz busca por nenhum tipo de carta além de Pokemons                                                            ####
#### Pokemons EX antigos deverão ser inseridos com -EX e Pokemons ex novos deverão ser inseridos sem o                ############################################################################################################################


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







#Lendo a tabela
tabela = pd.read_excel('BotTesteLiga\pteste.xlsx')
df = pd.DataFrame(tabela, columns=['Nome','Código', 'Coleção', 'Menor', 'Medio', 'Maximo'])

#Exibindo o conteúdo das cédulas 1 das colunas parametrizadas
print(tabela['Nome'].values[1] + ' ' + tabela['Código'].values[1])
#var = 'https://www.ligapokemon.com.br/?view=cards/card&card=' + tabela['Nome'].values[1]

#print('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[1]) + '%20(' + str(tabela['Código'].values[1][1:4]) + '/'+ str(tabela['Código'].values[1][5:7]) + ')&ed=' + str(tabela['Coleção'].values[1]) + '&num=' + str(tabela['Código'].values[1][1:4]))

#Quantidade de cartas 
n = int(input('Qual a quantidade de cartas para consulta?\n '))

#Abre o navegador Chrome atribuindo-o à variável
navegador = webdriver.Chrome()


def searchCard(qtd):
    for i in range(qtd):
        #if (tabela['Nome'].values[1][-2:] == 'ex'):
        #navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[i]) + '%20(' + str(tabela['Código'].values[i][1:4]) + '/'+ str(tabela['Código'].values[i][5:7]) + ')&ed=' + str(tabela['Coleção'].values[i]) + '&num=' + str(tabela['Código'].values[i]))

        navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[i]) + '%20' + str(tabela['Código'].values[i]) + '&ed=' + str(tabela['Coleção'].values[i]) + '&num=' + str(tabela['Código'].values[i]))
        #https://www.ligapokemon.com.br/?view=cards/card&card=Kindler%20(179/172)&ed=BRS&num=179
        #https://www.ligapokemon.com.br/?view=cards/card&card=Hoothoot%20(TG12/TG30)&ed=ARTG&num=TG12
        #else:
            #navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[1]) + '%20(' + str(tabela['Código'].values[1][1:4]) + '/'+ str(tabela['Código'].values[1][5:7]) + ')&ed=' + str(tabela['Coleção'].values[1]) + '&num=' + str(tabela['Código'].values[1]))
            #navegador.get('https://www.ligapokemon.com.br/?view=cards/card&card=' + str(tabela['Nome'].values[2]) + '%20(' + str(tabela['Código'].values[2][1:4]) + '/'+ str(tabela['Código'].values[2][5:7]) + ')&ed=' + str(tabela['Coleção'].values[2]) + '&num=' + str(tabela['Código'].values[2]))
        try:
            minVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[2]').get_attribute("textContent")
            midVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[4]').get_attribute("textContent")
            maxVal = navegador.find_element('xpath', '/html/body/main/div[4]/div[1]/div/div[3]/div[2]/div/div[6]').get_attribute("textContent")
            #print('Menor valor: ' + minVal + '\nValor médio: ' + midVal + '\nMaior Valor: ' + maxVal)
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]) , 'Menor']  = str(minVal)
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Medio'] = str(midVal)
            tabela.loc[tabela['Nome'] + tabela['Código'] + tabela['Coleção'] == str(tabela['Nome'].values[i]) + str(tabela['Código'].values[i]) + str(tabela['Coleção'].values[i]), 'Maximo'] = str(maxVal)
        except:
            print ('Pókemon: ' + str(tabela['Nome'].values[i] + str(tabela['Código'].values[i] + ' Não foi lido.')))
    return time.sleep(1)



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
