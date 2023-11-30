import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
# Descrever os passos manuais e depois transformar isso em código
webbrowser.open('https://web.whatsapp.com/')
sleep(5)
# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Plan1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome.capitalize()} estou testando um Bot, me de um retorno até a data {vencimento.strftime("%d/%m/%Y")}!'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        # Com base nos dados da planilha
        webbrowser.open(link_mensagem_whatsapp)
        sleep(5)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome.capitalize()}')
        with open('erros.csv', 'a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome.capitalize()},{telefone},')
