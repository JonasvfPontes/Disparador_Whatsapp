#bibliotecas necessárias
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
import urllib
import os
import sys
from tkinter import messagebox
from selenium.common.exceptions import NoSuchWindowException, NoSuchElementException
import funcoes


funcoes.mrkDiretory()
arquivo = 'Texto e Lista de contatos.xlsx'    

#Criar validador de usuário do disparo, pegando o nome do PC
df_usuário = pd.read_excel(arquivo, sheet_name='Config', usecols="A", engine='openpyxl')
nomeUsuario = df_usuário.loc[0, 'definicoes']
usuarioPCatual = os.getlogin() #Nome do Usuario/PC

if nomeUsuario == 'vazio' or nomeUsuario == '':
    #Adicionar nome de usuário ao Excel
    excel = load_workbook(arquivo)
    config = excel['Config']
    config['A2'].value = usuarioPCatual
    excel.save(arquivo)
elif nomeUsuario != usuarioPCatual: #Se o usuario não estover vazio e for diferente do nome existente Finalizar programa
    print('\nPermissão de usuário: NEGADA')
    messagebox.showinfo('Ação inválida', 
    'Esse PC não tem permissão para usar este programa. Para mais esclarecimetos entre em contato com o desenvolvedor (84) 98808-3657')
    funcoes.fSair()
print('\nPermissão de usuário: CONCEDIDA')

#Lendo opções de envio
df_opcoes = pd.read_excel(arquivo, sheet_name='Configurações', usecols="B:C", engine='openpyxl', header=None)
aguardar = int(df_opcoes.loc[1,2])
enviarImagem = True if str(df_opcoes.loc[0,2]).lower() == "sim" else False

#Lendo contatos
df_contatos = pd.read_excel(arquivo, sheet_name='Lista Contatos', usecols="A:C", engine='openpyxl')

#Lendo Mensagem e transformando em String
df_mensagem = pd.read_excel(arquivo, sheet_name='Texto', usecols="A", engine='openpyxl')
mensagem = str(df_mensagem.loc[0, 'Escreva o texto no campo abaixo']) #Pegando primeira linha da coluna 'Escrev o texto...'
print('Leitura do arquivo de contato e texto: Ok')

#Criando Objeto Navegador
print('Proxy = Desativado')
funcoes.fProxy(False) #Dasativando Proxy
print('Verificando atualizações do WebDriver')
try:
    servico = Service(ChromeDriverManager().install()) #Executando Serviço de atualização do WebDriver
    navegador = webdriver.Chrome(service=servico) #Criando objeto
    print('\nCriação do Objeto: Ok')
    funcoes.fProxy(True) #Ativando Proxy
    print('Proxy = Ativo')
    navegador.get('https://web.whatsapp.com/')
except Exception as e:
    funcoes.fProxy(True)
    print('Proxy = Ativo')
    exc_type, exc_obj, exc_tb = sys.exc_info()
    linha_erro = exc_tb.tb_lineno
    mensagem_erro = f"Erro na linha {linha_erro}:\nNome do erro: {e}"
    funcoes.fSair()

#Abrindo e aguardando login no WhatsApp
try:
    while len(navegador.find_elements(By.ID, 'side')) < 1:
        sleep(1)
except NoSuchWindowException:
    print('Navegador fechado manualmente')
    funcoes.fSair()
    
#Mandar mensaggem para cada contato -----------------------------------------------
sucesso = 0
insucesso = 0
try:
    for i, nome in enumerate(df_contatos['Primeiro Nome']):
        if len(str(df_contatos.loc[i, 'Numero'])) >= 12:
            try:
                numero = int(df_contatos.loc[i,'Numero'])
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
                        #Apertar Enter
                        navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                        #adicionar imagem---------------------------------
                        if enviarImagem:
                            funcoes.fEnviarImagem(navegador)
                        #-------------------------------------------------
                        sleep(aguardar) #Tempo depois que enviar cada mensagem
                        sucesso += 1
                        break
                    except NoSuchElementException:
                        #Esperar mais um pouco caso não encontrar XPATH
                        sleep(1)
                        cont = cont + 1
                        if cont >= 5:
                            insucesso += 1
            
            except Exception as e:
                exc_type, exc_obj, exc_tb = sys.exc_info()
                linha_erro = exc_tb.tb_lineno
                mensagem_erro = f"Erro na linha {linha_erro}: {e}"
                print('Erro na reprodução das mensagens.\nCausas possíveis:\n    •Navegador fechado manualmente\n    •Elemento no Navegador não encontrado')
                print("Nome da exceção:", type(e).__name__)
                print(mensagem_erro)
                funcoes.fSair()
except Exception as e:
    exc_type, exc_obj, exc_tb = sys.exc_info()
    linha_erro = exc_tb.tb_lineno
    mensagem_erro = f"Erro na linha {linha_erro}: {e}"
    print('''
Ocorreu um erro, talvez você tenha modificado o cabeçalho da lista de contatos
segue padrão da lista abaixo:
          Primeiro Nome     Numero      Adicional\n''')
    print("Nome da exceção:", type(e).__name__)
    print(mensagem_erro)
    funcoes.fSair()

navegador.close()
print('='*25, 'Fim da lista', '='*25)
print(f'''
Mensagens enviadas: {sucesso}
Numeros falhos:     {insucesso}
Total:              {sucesso + insucesso}''')
funcoes.fSair()
