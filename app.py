#bibliotecas necessárias
import pandas as pd
from openpyxl import workbook, load_workbook
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import urllib
import os
from tkinter import messagebox

#Criar validador de usuário do disparo, pegando o nome do PC
df_usuário = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Config', usecols="A", engine='openpyxl')

nomeUsuario = df_usuário.loc[0, 'definicoes']
usuarioPCatual = os.getlogin() #Nome do Usuario/PC

if nomeUsuario == 'vazio':
    #Adicionar nome de usuário ao Excel
    excel = load_workbook('Texto e Lista de contatos.xlsx')
    config = excel['Config']
    config['A2'].value = usuarioPCatual
    excel.save('Texto e Lista de contatos.xlsx')
elif nomeUsuario != usuarioPCatual: #Se o usuario não estover vazio e for diferente do nome existente Finalizar programa
    messagebox.showinfo('Ação inválida', 
    'Esse PC não tem permissão para usar este programa. Para mais esclarecimetos entre em contato com o desenvolvedor (84) 98808-3657')
    #('Ação inválida', 'Esse PC não tem permissão para usar este programa.')
    exit()
print('Permissão de usuário: CONCEDIDA')
#Lendo contatos

df_contatos = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Lista Contatos', usecols="A:B", engine='openpyxl')
#print(df_contatos)

#Lendo Mensagem e transformando em String
df_mensagem = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Texto', usecols="A", engine='openpyxl')
mensagem = str(df_mensagem.loc[0, 'Escreva o texto no campo abaixo']) #Pegando primeira linha da coluna 'Escrev o texto...'
print('Leitura do arquivo de contato e texto: Ok')

#Criando Objeto Navegador
print('\nVerificando atualizações do WebDriver\n')
servico = Service(ChromeDriverManager().install()) #Executando Serviço de atualização do WebDriver
navegador = webdriver.Chrome(service=servico) #Criando objeto

print('Criação do Objeto: Ok')
navegador.get('https://web.whatsapp.com/')

#Abrindo e aguardando login no WhatsApp
while len(navegador.find_elements(By.ID, 'side')) < 1:
    sleep(1)

#Mandar mensaggem para cada contato
for i, nome in enumerate(df_contatos['Primeiro Nome']):
    if len(str(df_contatos.loc[i, 'Numero'])) >= 12:
        try:
            numero = str(int(df_contatos.loc[i,'Numero']))
            if str(nome) == 'nan':
                mensagemPersonalizada = urllib.parse.quote(mensagem.replace('--CONTATO--', ''))
            else:
                mensagemPersonalizada = urllib.parse.quote(mensagem.replace('--CONTATO--', str(nome)))
            link = f'https://web.whatsapp.com/send?phone={numero}&text={mensagemPersonalizada}'
            navegador.get(link)
            while len(navegador.find_elements(By.ID, 'side')) < 1:
                sleep(2)
            sleep(5) #Após carregar ID side, esperar 5 segundos
            #Procurar  xPATH (Enter), se error esperar até 5 segundos, senão ir para o próximo contato
            cont=0 #Contador de espera
            while cont <5:
                try:
                    navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                    sleep(15) #Tempo depois que enviar cada mensagem
                    #adicionar imagem---------------------------------

                    #-------------------------------------------------
                    break
                except:
                    #Esperar mais um pouco caso não encontrar XPATH
                    sleep(2)
                    cont = cont + 1
           
        except:
            print('Erro na reprodução das mensagens.\nCausas possíveis:\nNavegador fechado manualmente\n\n')
            sleep(5)
            exit()

print('Fim da lista\n')
navegador.close