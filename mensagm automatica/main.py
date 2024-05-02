
# Desenvolver um boot para automatizar mensagens no WhatssApp Web

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(5)

workbook = openpyxl.load_workbook('contatos.xlsx')
pagina_contatos = workbook['Planilha1']

for linha in pagina_contatos.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    prazo = linha[2].value
    mensagem = f'Olá {nome}, gostaria que você escrevesse uma carta especial para a Nathália pois ela vai ao EJC. Se possível me envie até o dia {prazo.strftime('%d/%m/%Y')}'
    
    try:
    
        link = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link)
        sleep(10)
        pyautogui.press('enter')
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)

    except:
        
        print(f'Não foi possível enviar mensagem para a pessoa {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as file:
            file.write(f'{nome}, {telefone}')



