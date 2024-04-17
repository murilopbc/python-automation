from PySimpleGUI import PySimpleGUI as sg

sg.theme('Raddit')
layout = [
    [sg.Text('Usuário'), sg.Input(key='usuario')],
    [sg.Text('Senha'), sg.Input(key='senha', password_char='*')],
    [sg.Button('Entrar')]


]

janela = sg.Window('Tela de Login', layout)

while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == 'Entrar':
        if valores ['usuario'] == 'murilo' and valores['senha'] == '123':
            print('Seja bem-vindo!')
