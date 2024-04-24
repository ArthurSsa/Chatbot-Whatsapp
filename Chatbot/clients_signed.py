import openpyxl
import webbrowser
from urllib.parse import quote
from time import sleep
import pyautogui
import datetime

def enviando_mensagem(mensagem:str, telefone:int):
    """Manda a mensagem para o número de telefone do cliente

    Args:
        mensagem (str): mensagem personalizada
        telefone (int): telefone do cliente
    """
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)
        click = pyautogui.locateCenterOnScreen('seta.png')
        sleep(3)
        pyautogui.click(click[0], click[1])
        sleep(3)
        pyautogui.hotkey('ctrl', 'w')
        sleep(3)
    except:
        print("Algo deu errado ao enviar a mensagem...")
        with open("erros.csv", "a",newline='', encoding='utf-8') as arquivo:
            arquivo.write(f"{telefone}")


def identificar_periodo(nome: str, telefone: int, periodo: str):
    """Identifica o período do dia e o período de maior atividade do usuário

    Args:
        nome (str): Nome do cliente
        telefone (int): Telefone do cliente
        periodo (str): Período de maior atividade
    """
    hora_atual = datetime.datetime.now().time()

    match periodo:
        case "Manhã":
            if hora_atual.hour<12:
                mensagem = f"Bom dia flor do dia, confira as ofertas dessa manhã {nome}!"
                enviando_mensagem(mensagem, telefone)
        case "Tarde":
            if hora_atual.hour in range(12, 19):
                mensagem = f"Boa tarde {nome}, está na hora de fazer uma merenda!"
                enviando_mensagem(mensagem, telefone)
        case "Noite":
            if hora_atual.hour>18 :
                mensagem = f"A noite só é boa com uma boa janta, aproveite {nome}!"
                enviando_mensagem(mensagem, telefone)



def identificando_clientes():
    """Percorre a lista de clientes cadastrados
    """
    clientes_cadastrados = openpyxl.load_workbook("clientes_cadastrados.xlsx")
    página_cadastros = clientes_cadastrados["Perfis"]

    print("início")

    for linha in página_cadastros.iter_rows(min_row=2):
        nome = linha[0].value #Nome do usuário
        telefone = int(linha[1].value) # telefone para acessar
        periodo = linha[2].value #manhã, tarde ou noite
        try:
            identificar_periodo(nome, telefone, periodo)
        except:
            print("Algo deu errado ao identificar os clientes")


identificando_clientes()