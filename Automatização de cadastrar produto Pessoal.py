import pyautogui
import time

pyautogui.PAUSE = 0.3

pyautogui.press("win")
time.sleep(3)
pyautogui.write("edge")
pyautogui.press("enter")
time.sleep(3)
pyautogui.press("Ctrl+l")
time.sleep(3)
pyautogui.press("Crlt+l")
time.sleep(5)
pyautogui.write("https://dlp.hashtagtreinamentos.com/python/intensivao/login")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=724, y=352)
pyautogui.write("pythonimpressionador@gmail.com")
pyautogui.press("tab")
pyautogui.write("sua senha")
pyautogui.press("tab")
pyautogui.press("enter")
time.sleep(3)
pyautogui.click(x=483, y=510)
time.sleep(3)
pyautogui.click(x=287, y=243)


import pandas as pd

tabela = pd.read_excel("produtospapelaria.xlsx")
print(tabela)

for linha in tabela.index:
    
    codigo = tabela.loc[linha, "Unnamed: 0"]
    pyautogui.write(str(codigo))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "Unnamed: 16"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "Unnamed: 5"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "Unnamed: 3"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "Unnamed: 8"]))
    pyautogui.press("tab")
    pyautogui.write(str(tabela.loc[linha, "Unnamed: 7"]))
    pyautogui.press("tab")
    obs = tabela.loc[linha, "Unnamed: 10"]
    if not pd.isna(obs):
        pyautogui.write(str(tabela.loc[linha, "Unnamed: 10"]))
    pyautogui.press("tab")
    pyautogui.press("enter")
    pyautogui.click(x=466, y=247)