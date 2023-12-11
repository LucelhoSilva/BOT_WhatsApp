import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value

    mensagem = f'Olá {nome} tudo bem ?. 🚀 Venha fazer parte dessa comunidade no Discord !! 🚀🔗 Link de Convite: https://discord.com/invite/QXzdGW8FaT Aqui você vai encontrar: Vagas de emprego 💼 , Freelancer 🛠e Cursos gratuitos 📚 disponíveis 24 horas. Além de espaço para desenvolver projetos colaborativos 🤝 efazer networking 🌐'

    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(20)  
        seta = pyautogui.locateCenterOnScreen(
            'seta.png', grayscale=True, confidence=0.8)
        if seta is not None:
            pyautogui.click(seta[0], seta[1])
            sleep(5)  
            pyautogui.hotkey('ctrl', 'w')
        else:
            print(f'Botão de envio não encontrado para {nome}')
        sleep(5)  
    except Exception as e:
        print(f'Não foi possível enviar mensagem para {nome}, erro: {e}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
