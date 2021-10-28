#!/usr/bin/env python
# coding: utf-8

# instalar as libs abaixo n seu projeto com os comandos:
# pip install pyautogui
# pip install pyperclip

# importe as bibliotecas que iremos usar neste projeto
import pyautogui
import pyperclip
import time

# Passo 1: Entrar no sistema (no caso, entrar no link)
# aqui partimos do pressuposto que o navegador já se encontra aberto!

pyautogui.PAUSE = 2 # pausa de 2 segundos para executar as ações do pyautogui
pyautogui.hotkey("ctrl", "t") # pressionar as teclas ctrl + t para abrir uma nova aba do navegador

# o pyautogui não consegue escrever links com alguns caracteres, no exemplo abaixo seria a ?
# então usarei o pyperclip para copiar o link abaixo
pyperclip.copy("https://drive.google.com/drive/folders/1o8MdkzoZkVS68dlTMI7w1NYcjcvPsK-b?usp=sharing")

pyautogui.hotkey("ctrl", "v") # pressionaremos o ctrl + v para colar o link na barra de endereços do navegador
pyautogui.press("enter") # pressionar enteder para confirmar a ação

# Passo 2: Navegar até o local do arquivo (entrar na pasta exportar)
time.sleep(5) # aguardar 5 segundos para a tela ser carregada completamente

## pyautogui.position() # comando usado para capturar a posição do mouse na tela
pyautogui.click(x=332, y=303, clicks=2) # comando para executar o click do mouse na posição da tela que desejamos

# Passo 3: Fazer o download do relatorio
# os clicks abaixo, se referem ao clique no relatorio / clique nos pontinhos de opções / clique em Download para baixar
pyautogui.click(x=392, y=367)
pyautogui.click(x=1159, y=196)
pyautogui.click(x=978, y=585)

time.sleep(10) # pausa de 10 segundos para o navegador abrir a tela de salvar o arquivo
pyautogui.press("enter") # confirmar com o enter
time.sleep(5) # aguardar 5 segundos para baixar o arquivo


# Passo 4: Calcular os indicadores
# instale a lib do pandas com o comando abaixo:
# pip install pandas

import pandas as pd # importar o pandas para trabalhar com os dados

## usa o r no inicio do link do arquivo para que o python entenda que tem que usar os caracteres completos
tabela = pd.read_excel(r"E://BackupPC/Downloads/Vendas - Dez.xlsx")
display(tabela) # mostrar tabela (display para o jupiter, na sua IDE seria o print)

# criamos as variaveis armazenando a soma das colunas valor final e quantidade
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 5: Entrar no e-mail
pyautogui.hotkey("ctrl", "t") # comando para abrir outra aba do navegador
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox") # copiar o link do email
pyautogui.hotkey("ctrl", "v") # cola o link no navegador
pyautogui.press("enter") # confirma o navegador usar o link que colamos

time.sleep(10) #aguarda 10 segundos para dar o proximo comando

# Passo 6: Enviar por e-mail o resultado
pyautogui.click(x=98, y=195) #clica no botao de escrever email

pyautogui.write("seu-email-aqui@gmail.com") #escreve o email

pyautogui.press("tab") #seleciona o e-mail

time.sleep(2) #aguarda 2 segundos para o proximo comando

pyautogui.press("tab") #avança para o campo de assunto do email

pyperclip.copy("Relatório de Vendas")#copia o assunto pra area de transferencia

pyautogui.hotkey("ctrl", "v") #cola o assunto no campo de assunto do e-mail

pyautogui.press("tab") #pula pro corpo do e-mail

# criar o texto que irá no e-mail e armazenar numa variavel
texto = f"""
Prezados, bom dia!

O Faturamento ontem foi de: R$ {faturamento:,.2f}
A Quantidade de produtos vendidas foi de: {quantidade:,}

Att,

Adériton Prado
"""

pyperclip.copy(texto) #copia o texto que está na variavel

pyautogui.hotkey("ctrl", "v") # cola o texto no e-mail

pyautogui.hotkey("ctrl", "enter") #envia o e-mail



