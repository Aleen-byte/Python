
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui


webbrowser.open('https://web.whatsapp.com')
sleep(10)
pyautogui.hotkey('ctrl', 'w')

workbook = openpyxl.load_workbook('clientes.xlsx')

pagina_clientes = workbook['Planilha1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    valor = linha[2].value
    vencimento = linha[3].value

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. por favor pague logo seu caloteiro!'


    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        pyautogui.hotkey('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possível enviar a  para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')