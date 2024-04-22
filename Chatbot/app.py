import openpyxl
import webbrowser
from urllib.parse import quote
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(5)

#ler a planilha e guardar informação
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Contatos']

for linha in pagina_clientes.iter_rows(min_row=2):
    #variaveis de nome, telefone, imagem e mensagem
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = f'olá {nome} aqui está o nosso cardápio do dia:\nBife acebolado \nCostelinha suína acebolada \nFilé de peito de frango \nFrango assado ao forno \nCreme de galinha \nLinguiça '

    #link personalizado
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(5)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(1)
        pyautogui.click(seta.x,seta.y)
        sleep(1)
        pyautogui.hotkey('ctrl','w')
        sleep(1)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
           arquivo.write(f'{nome},{telefone}')