from time import sleep
from urllib.parse import quote
import webbrowser
import pyautogui
import openpyxl

webbrowser.open("https://web.whatsapp.com")
sleep(30)

workbook = openpyxl.load_workbook("Relatório.xlsx")


pagina_clientes = workbook["Planilha1"]


for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    print(nome)
    print(telefone)
    mensagem_whatsapp = f'Olá, {
        nome}!\n Sou um Robozinho 🤖 e queria apenas te lembrar de enviar o relatório'

    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
        telefone}&text={quote(mensagem_whatsapp)}'

    webbrowser.open(link_mensagem_whatsapp)
    sleep(20)

    try:
        seta = pyautogui.locateCenterOnScreen("./imagens/seta.PNG")
        sleep(10)

        pyautogui.click(seta[0], seta[1])
        sleep(5)

        pyautogui.hotkey("ctrl", "w")
        sleep(5)

    except:
        print(f"Não foi possível enviar mensagem para: {nome}")
        with open("erros.csv", "a", newline="", encoding="utf-8") as arquivo:
            arquivo.write(f"{nome}, {telefone}")
