import pyautogui
import time
import pandas as pd 

x = 0
y = 0 

caminho = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/ASNDV/DEVOLUÇÃO GELADEIRAS.xlsx"
ler = pd.read_excel(caminho)
tabela = pd.DataFrame(ler)
print(tabela)

while True and x < len (tabela):

    ilpn = tabela.iloc[x, y + 2]
    plu = tabela.iloc[x, y]
    quantidade = tabela.iloc[x, y + 3]
    dia_validade = tabela.iloc[x, y + 4 ]
    mes_validade = tabela.iloc[x, y + 5 ]
    ano_validade = tabela.iloc[x, y + 6 ]
    dia_fabricacao = tabela.iloc[x, y + 7]
    mes_fabricacao = tabela.iloc[x, y + 8]
    ano_fabricacao = tabela.iloc[x, y + 9]

    print (plu)
    print(ilpn)
    time.sleep(0.5)
    pyautogui.doubleClick(131, 254)
    time.sleep(0.5)
    pyautogui.write("0")
    time.sleep(1)
    pyautogui.write(str(ilpn))
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(2.2)
    pyautogui.hotkey("ctrl" , "a")
    time.sleep(2)
    pyautogui.hotkey("ctrl" ,"a")
    time.sleep(2)
    pyautogui.leftClick(65, 238)
    time.sleep(1)
    pyautogui.leftClick(65, 238)
    time.sleep(2)
    pyautogui.write(str(plu))
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.doubleClick(50, 244)
    time.sleep(2)
    pyautogui.write(str(quantidade))
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.leftClick(97, 220)
    pyautogui.write(str(dia_validade))
    time.sleep(1)
    pyautogui.leftClick(155, 220)
    pyautogui.write(str(mes_validade))
    time.sleep(1)
    pyautogui.leftClick(218, 220)
    pyautogui.write(str(ano_validade))
    time.sleep(1)
    pyautogui.leftClick(99, 241)
    time.sleep(1)
    pyautogui.leftClick(99, 241)
    pyautogui.write(str(dia_fabricacao))
    time.sleep(1)
    pyautogui.leftClick(153, 241)
    time.sleep(1)
    pyautogui.leftClick(153, 241)
    pyautogui.write(str(mes_fabricacao))
    time.sleep(1)
    pyautogui.leftClick(219, 243)
    time.sleep(1)
    pyautogui.leftClick(219, 243)
    pyautogui.write(str(ano_fabricacao))
    time.sleep(2)
    print(dia_fabricacao)
    pyautogui.leftClick(421, 209)
    time.sleep(2)
    x = x + 1