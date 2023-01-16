import pyautogui
import pyperclip
import time
import pandas as pd

# Passo 1: Entrar no sistema da empresa  (link no drive)
pyautogui.PAUSE = 1



pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

# Passo 2: Navegar ate o local do relatorio  (entrar na pasta Exportar)
time.sleep(5)
pyautogui.click(x=499, y=323, clicks=2)


#Passo 3: Exportar o relatorio (fazer o download)
time.sleep(1.5)
pyautogui.click(x=500, y=319)
pyautogui.click(x=1660, y=187)
pyautogui.click(x=1309, y=693)
time.sleep(3)
pyautogui.click(x=1710, y=1036)


#Passo 4: calcular os indicadores (faturamento e quantidade de produtos)

tabela = pd.read_excel(r"C:\Users\Dante\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()                                                     
quantidade = tabela["Quantidade"].sum()
print(faturamento)
print(quantidade)

#Passo 5: Enviar um e-mail para a diretoria

##abrir aba e entrar no email
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/1/?ogbl#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

time.sleep(5)

##clicar no botao escrever
pyautogui.click(x=228, y=181)
time.sleep(10)

##preencher as informacoes do email
#destinatario
pyautogui.write("sombradante02@gmail.com")
pyautogui.press("tab",presses=2)

#assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
                         
#corpo
texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f} 
A quantidade de produtos foi de: {quantidade:,}

Abraços
Dante"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

## enviar o email
pyautogui.hotkey("ctrl", "enter")

