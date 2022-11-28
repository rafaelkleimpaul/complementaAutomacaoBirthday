import PySimpleGUI as sg
import pandas as pd

# Adiciona cores a janela
sg.theme('DarkTeal9')

# Definindo o arquivo xls
EXCEL_FILE = 'dados.xlsx'
df = pd.read_excel(EXCEL_FILE)

# Campos a sere adicionados ao form
layout = [[sg.Text('Preencha os seguintes campos: ')],
          [sg.Text('Nome', size=(15, 1)),
           sg.InputText(key='Nome')],
          [sg.Text('Sobrenome', size=(15, 1)),
           sg.InputText(key='Sobrenome')],
          [sg.Text('Data de hoje', size=(15, 1)),
           sg.InputText(key='Data')],
          [
              sg.Text('Mensagem', size=(15, 1)),
              sg.Multiline(size=(42, 10), key='Mensagem')
          ], [sg.Submit(), sg.Exit()]]

window = sg.Window('Entrada de informacoes', layout)


# Limpa Campos
def clear_inputs():
    for key in values:
        window['Nome'].update(''),
        window['Sobrenome'].update(''),
        window['Mensagem'].update('')


while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Submit':
        df = df.df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Dados salvos')
        clear_inputs()

window.close()