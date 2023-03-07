import pyautogui
import pyperclip
import time 
import pandas as pd


#abri o navegador
pyautogui.alert("A automação ira começar\nprecione ok")
pyautogui.press("winleft")
pyautogui.write("Chrome")
time.sleep(5)
pyautogui.press("enter")
time.sleep(5)

#entra no drive
pyautogui.click(364, 50, clicks=1)
time.sleep(2)
link = "https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(25)

# download do arquivo 
pyautogui.click(357, 345, clicks=1)
time.sleep(10)
pyautogui.click(1159,182, clicks=1)
time.sleep(10)
pyautogui.click(932, 624, clicks=1)
time.sleep(30)
pyautogui.click(671, 507, clicks=1)

#mexendo no arquivo 
time.sleep(20)
df = pd.read_excel(r"C:\Users\Cliente\Downloads\Vendas - Dez.xlsx")
faturamento = df['Valor Final'].sum()
quat_prod =df['Quantidade'].sum()

# resultado
pyautogui.alert(f"O faturamento foi R${faturamento:,.2f} \ne a quantidade vendida {quat_prod}\nprecione ok")
