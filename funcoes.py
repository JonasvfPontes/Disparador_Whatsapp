from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import os
from pathlib import Path
from selenium.common.exceptions import NoSuchElementException
import subprocess
import sys

# Função para enviar imagens
def fEnviarImagem(navegador):
    # Pega o caminho da pasta completa do diretório
    caminhoMidia = Path().absolute()
    caminhoMidia = str(caminhoMidia) + str('\\Imagens')

    for nome_arquivo in os.listdir(caminhoMidia):
        try:
            if nome_arquivo.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.webp')):

                sleep(1)
                # Localiza o ícone de anexo e clique nele
                botaoMais = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/div')
                                                                
                # Construir caminho da imagem
                imagemAtual = os.path.join(caminhoMidia, nome_arquivo)

                # localiza botãoGaleria dentro do botãoMais
                botaoGaleria =  botaoMais.find_element(By.XPATH,'/html/body/div[1]/div/div/div[5]/div/footer/div[1]/div/span[2]/div/div[1]/div/div/span/div/ul/div/div[2]/li/div/input')
                botaoGaleria.send_keys(imagemAtual)
                sleep(1)  # Aguardar imagem ser carregada

                # Enviar imagem
                navegador.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').send_keys('\n')
                sleep(1)  # Aguardar imagem ser enviada
        except NoSuchElementException:
            print('Não consegui enviar imagem')


#verificando se arquivo .bat existe
def fProxy(ativar):
    '''Se "ativar = True então ProxyEnable = 1
       Se "ativar = False então ProxyEnable = 0'''
    
    if not os.path.exists(r"C:\automacao\proxy0-Whatsapp.vbs"):
        # Se o arquivo não existir, crie-o
        with open(r"C:\automacao\proxy0-Whatsapp.vbs", "w") as arquivo:
            arquivo.write('dim oShell\n')
            arquivo.write('set oShell = Wscript.CreateObject("Wscript.Shell")\n\n')
            arquivo.write('oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"\n\n')
            arquivo.write('Set oShell = Nothing\n')
            arquivo.write("'Desativa o proxy")

    if not os.path.exists(r"C:\automacao\proxy1-Whatsapp.vbs"):
        # Se o arquivo não existir, crie-o
        with open(r"C:\automacao\proxy1-Whatsapp.vbs", "w") as arquivo:
            arquivo.write('dim oShell\n')
            arquivo.write('set oShell = Wscript.CreateObject("Wscript.Shell")\n\n')
            arquivo.write('oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"\n\n')
            arquivo.write('Set oShell = Nothing\n\n')
            arquivo.write("'Ativa o proxy")

    if ativar:
        # Execute o arquivo .bat
        # Caminho para o arquivo .vbs que você deseja executar
        caminho_arquivo_vbs = r'C:\automacao\proxy1-Whatsapp.vbs'
    else:
        caminho_arquivo_vbs = r'C:\automacao\proxy0-Whatsapp.vbs'
    
    # Execute o arquivo .vbs
    subprocess.run(['wscript.exe', caminho_arquivo_vbs], capture_output=True, text=True)


def fSair():
    print('\n', '='*50)
    print('Aperte "Enter" para sair')
    input()
    print('Saindo...')
    print('Feche essa tela')
    try:
        sys.exit()
    except:
        print()

def mrkDiretory():
    diretorio1 = r"C:\automacao"
    if not os.path.exists(diretorio1):
        os.makedirs(diretorio1)

    diretorio2 = os.path.abspath('.\\Imagens')
    if not os.path.exists(diretorio2):
        os.makedirs(diretorio2)