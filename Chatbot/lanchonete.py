import openpyxl
import webbrowser
from urllib.parse import quote
from time import sleep
import pyautogui
import datetime
from clients_signed import identificar_periodo

#Inicia o whatsapp web
webbrowser.open('https://web.whatsapp.com/')
sleep(10)

#ler a planilha e guardar informação
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Contatos']

for linha in pagina_clientes.iter_rows(min_row=2):
    #variaveis de nome, telefone, imagem e mensagem
    nome = linha[0].value
    telefone = linha[1].value
    turno = linha[2].value

    if turno == 'noite':
        mensagem = f'Olá {nome} boa noite, seja bem vindo a Huawei Lanches, segue em anexo nosso cardápio de delícias '
        mensagem2 = f'esperamos que o seu jantar seja maravilhoso '
    elif turno == 'dia':
        mensagem = f'Olá {nome} bom dia, seja bem vindo a Huawei Lanches, segue em anexo nosso cardápio de delícias '
        mensagem2 = f'esperamos que tenha um delicioso almoco '
    #link personalizado
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(7)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(1)
        pyautogui.click(seta.x,seta.y)
        sleep(1)
        mais = pyautogui.locateCenterOnScreen('mais.png')
        sleep(1)
        pyautogui.click(mais.x,mais.y)
        sleep(1)
        documento = pyautogui.locateCenterOnScreen('doc.png')
        sleep(1)
        pyautogui.click(documento.x,documento.y)
        sleep(1)
        pyautogui.write(r"C:\Users\arthu\Downloads\cardapio.pdf")
        pyautogui.press('enter')
        sleep(1)
        pyautogui.press('enter')
        sleep(1)
        pyautogui.write((mensagem2))
        sleep(3)
        pyautogui.press('enter')
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
        identificar_periodo(nome, telefone, turno)
        
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
           arquivo.write(f'{nome},{telefone}')