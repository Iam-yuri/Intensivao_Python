import pyautogui
import pyperclip
import time

from Tools.scripts.dutree import display

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(10)

pyautogui.click(x=974, y=676, clicks=2)
time.sleep(10)

pyautogui.click(x=922, y=828)
pyautogui.click(x=3339, y=404)
pyautogui.click(x=2890, y=1406)

time.sleep(10)

import pandas as pd

tabela = pd.read_excel(r"C:\Users\Yu ri\Downloads\Vendas - Dez")
display(tabela)

faturamento = tabela["Valor Final"].sum()
qtde_produtos = tabela["Quantidade"].sum()

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/?ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(10)

pyautogui.click(x=282, y=463)

pyautogui.write("iamyuri2003@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")

pyautogui.write("Relatorio de vendas")
pyautogui.press("tab")

texto = f"""
Prezados, bom dia!

O faturamento de ontem foi de R$: faturamento
A quantidade de produtos foi de: quantidade

Ats
Yuri
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl", "enter")