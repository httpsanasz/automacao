# automacao

automação em python

import pyautogui

import pyperclip

import time

#Passo 1: Entrar no sistema da empresa

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl","t")

pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")

pyautogui.hotkey("ctrl","v")

pyautogui.press("enter")#Passo 2: Navegar no sistema e encontra a base de vendas

time.sleep(3)

pyautogui.position() 

pyautogui.click(x=362, y=304, clicks= 2)#Passo 3: Fazer o download da base de vendas

time.sleep(3)

pyautogui.position()

pyautogui.click(x=193, y=315)

time.sleep(3)

pyautogui.position()

pyautogui.click(x=1291, y=92)#Passo 4: Importar base de dados para o Python

import pandas

tabela = pandas.read_excel(r"C:\Users\User\Downloads\Vendas - Dez.xlsx")

display(tabela)

#Passo 5: Calcular os indicadores da empresa

faturamento = tabela["Valor Final"].sum()

print(faturamento)

quantidade = tabela["Quantidade"].sum()

print(quantidade)

#Passo 6: Enviar um e-mail para diretoria com os indicadores de vendas

pyautogui.PAUSE = 2

pyautogui.hotkey("ctrl","t")

pyperclip.copy(r"https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")

pyautogui.hotkey("ctrl","v")

pyautogui.press("enter")

time.sleep(8)

pyautogui.position()

pyautogui.click(x=45,y=96)

pyautogui.click(x=52,y=167)

pyautogui.write("anaohi.1420@aluno.saojudas.br")

pyautogui.press("tab")

pyautogui.press("tab")

pyautogui.write("Relatório de vendas")

pyautogui.press("tab")

pyautogui.write(f""" 
Relatório de vendas 
R${faturamento :,.2f}
{quantidade}
""")

pyautogui.hotkey("ctrl","enter")
