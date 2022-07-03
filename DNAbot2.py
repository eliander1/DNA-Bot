import PyPDF2
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
import urllib
from tkinter import *


def AtualizarDevedores():
    #  ABRINDO ARQUIVO  PDF EM MODO LEITURA E LENDO O BINARIO
    devPdf = open(r'PlanilhaSistemaConvertida.pdf', 'rb')
    pdfDevedores = PyPDF2.PdfFileReader(devPdf)  # APÓS PEGAR O BINARIO LEMOS O DADO DO PDF PELO BINARIO
    tamanho = pdfDevedores.numPages  # numero de páginas no pdf

    bancoDados = pd.read_excel('BancoDadosGeralDNA.xlsx')
    listaDevTel = []
    listaDevNome = []
    listaMsg = []
    i = 0
    check = []
    geraldf = pd.DataFrame(bancoDados['NOME'], columns=['NOME'])

    while i < tamanho:
        pagina = pdfDevedores.getPage(i)
        try:
            textopagina = pagina.extractText()
            textopagina = re.sub('\n', '', textopagina)
            textopagina = re.sub(' ', '', textopagina)

            for item, cel in enumerate(bancoDados['CELULAR']):
                if cel in textopagina:
                    if cel not in listaDevTel:
                        listaDevTel.append(cel)
                        listaDevNome.append(geraldf.at[item, 'NOME'])
                        listaMsg.append("")
                        check.append("")
        except:
            print(f"página {i}")
        i += 1

    print(len(listaDevNome), len(listaDevTel), len(listaMsg))
    print(listaDevNome, listaDevTel, listaMsg)
    df = pd.DataFrame({'Nome': listaDevNome, 'Telefone': listaDevTel, 'Mensagem': listaMsg, 'Check': check})
    df.to_excel(r"CobrancasDoDia.xlsx", index=False)


def cobrarGeral():
    # importando a planilha excel
    listacontatos = pd.read_excel(r"CobrancasDoDia.xlsx", engine='openpyxl')

    # fazendo o chrome abrir o whatsapp web
    s = Service('chromedriver.exe')
    navegador = webdriver.Chrome(service=s)
    navegador.get("https://web.whatsapp.com/")

    # a cada 10 segundos confira se foi logado
    while len(navegador.find_elements(By.XPATH, "//*[@id=\"pane-side\"]")) < 1:
        sleep(10)

    tempoEnvio = 7
    tempoEspera = 7

    # Quando logado, envie a mensagem da planilha para cada contato dela
    for i, mensagem in enumerate(listacontatos['Mensagem']):
        if listacontatos.loc[i, 'Check'] != 'ok':

            try:
                Telefone = listacontatos.loc[i, "Telefone"]
                Telefone = re.sub(r'[^\w]', ' ', Telefone)
                Telefone = re.sub(' ', '', Telefone)
                texto = urllib.parse.quote(f"{mensagem}")
                link = f"https://web.whatsapp.com/send?phone=55{Telefone}&text={texto}"
                navegador.get(link)

                # esperar o whatsapp com a msg do contato carregar
                while len(navegador.find_elements(By.XPATH, "//*[@id=\"pane-side\"]")) < 1:
                    sleep(tempoEnvio)

                # pressionar o enter e esperar 7 segundos pra iniciar o loop de novo
                navegador.find_element(By.XPATH, "//*[@id=\"main\"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button").send_keys(Keys.ENTER)
                listacontatos.loc[i, 'Check'] = 'ok'
                sleep(tempoEspera)

            except:
                print("erro")
                listacontatos.loc[i, 'Check'] = 'erro'

    listacontatos.to_excel(r"CobrancasDoDia.xlsx", index=False)
    print(listacontatos)


# INTERFACE

janela = Tk()
janela.title("Bot DNA")

texto_orientacao = Label(janela, text='Clique no botão abaixo para gerar uma planilha com todos os devedores')
texto_orientacao.grid(column=0, row=0, padx=10, pady=10)

botao = Button(janela, text='Gerar Planilha de Cobrança', command=AtualizarDevedores, bg="green", fg="white")
botao.grid(column=0, row=1, padx=10, pady=10)

texto_orientacao = Label(janela, text='Clique no botão abaixo para cobrar todos os devedores da planilha gerada acima')
texto_orientacao.grid(column=0, row=2, padx=10, pady=10)

botao = Button(janela, text='Cobrar Todo Mundo', command=cobrarGeral, bg="green", fg="white")
botao.grid(column=0, row=3, padx=10, pady=10)

# andamento = Label(janela, text='0 de ')
# andamento.grid(column=0, row=4, padx=10, pady=10)

# sempre colocar isso no final pra manter a janela aberta
janela.mainloop()
