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
from selenium.common.exceptions import NoSuchWindowException

def fSair():
    print('\n', '='*50)
    print('Aperte "Enter" para sair')
    input()
    print('Saindo...')
    exit()
        

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
    fSair()
print('\nPermissão de usuário: CONCEDIDA')

#Lendo opções de envio
df_opcoes = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Configurações', usecols="B:C", engine='openpyxl', header=None)
aguardar = int(df_opcoes.loc[1,2])


#Lendo contatos
df_contatos = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Lista Contatos', usecols="A:C", engine='openpyxl')

#Lendo Mensagem e transformando em String
df_mensagem = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Texto', usecols="A", engine='openpyxl')
mensagem = str(df_mensagem.loc[0, 'Escreva o texto no campo abaixo']) #Pegando primeira linha da coluna 'Escrev o texto...'
print('Leitura do arquivo de contato e texto: Ok')

#Criando Objeto Navegador
print('Verificando atualizações do WebDriver')
servico = Service(ChromeDriverManager().install()) #Executando Serviço de atualização do WebDriver
navegador = webdriver.Chrome(service=servico) #Criando objeto
print('\nCriação do Objeto: Ok')
navegador.get('https://web.whatsapp.com/')

#Abrindo e aguardando login no WhatsApp
try:
    while len(navegador.find_elements(By.ID, 'side')) < 1:
        sleep(1)
except NoSuchWindowException:
    print('Navegador fechado manualmente')
    fSair()
    

#Mandar mensaggem para cada contato -----------------------------------------------
sucesso = 0
insucesso = 0
try:
    for i, nome in enumerate(df_contatos['Primeiro Nome']):
        if len(str(df_contatos.loc[i, 'Numero'])) >= 12:
            try:
                numero = df_contatos.loc[i,'Numero']
                adicional = df_contatos.loc[i, 'Adicional']
                
                #Personalizando mensagem
                if pd.isna(adicional): #Verifica se adicional está vazio
                    mensagemPersonalizada = mensagem.replace('--ADD--', '')
                else:
                    mensagemPersonalizada = mensagem.replace('--ADD--', str(adicional))

                if pd.isna(nome): #Verifica se nome está vazio
                    mensagemPersonalizada = urllib.parse.quote(mensagemPersonalizada.replace('--CONTATO--', ''))
                else:
                    mensagemPersonalizada = urllib.parse.quote(mensagemPersonalizada.replace('--CONTATO--', str(nome)))

                #Enviando mensagem
                link = f'https://web.whatsapp.com/send?phone={numero}&text={mensagemPersonalizada}'
                navegador.get(link)
                while len(navegador.find_elements(By.ID, 'side')) < 1: #Confere se a página carregou
                    sleep(1)
                sleep(5) #Após carregar ID side, esperar alguns segundos só pra assegurar que a página irá carregar
                
                #Procurar  xPATH (Enter), se error esperar até 5 segundos, senão encontrar
                #significa quer certamente o numero não existe, então devo ir para o próximo contato
                cont=0 #Contador de espera
                while cont <5:
                    try:
                        navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                        sleep(aguardar) #Tempo depois que enviar cada mensagem
                        #adicionar imagem---------------------------------

                        #-------------------------------------------------
                        sucesso += 1
                        break
                    except:
                        #Esperar mais um pouco caso não encontrar XPATH
                        sleep(1)
                        cont = cont + 1
                        if cont >= 5:
                            insucesso += 1
            
            except Exception as e:
                print('Erro na reprodução das mensagens.\nCausas possíveis:\nNavegador fechado manualmente')
                print("Nome da exceção:", type(e).__name__)
                fSair()
except Exception as e:
    print('''
Ocorreu um erro, talvez você tenha modificado o cabeçalho da lista de contatos
segue padrão da lista abaixo:
          Primeiro Nome     Numero      Adicional\n''')
    
    print("Nome da exceção:", type(e).__name__)
    fSair()

navegador.close
print('='*25, 'Fim da lista', '='*25)
print(f'''
Mensagens enviadas: {sucesso}
Numeros falhos:     {insucesso}
Total:              {sucesso + insucesso}''')
fSair()
