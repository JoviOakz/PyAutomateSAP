import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 0.05

bot.click(1802, 14)

excel_path = "LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_tab(x):
    for _ in range(x):
        bot.press('tab')

def press_shtab(x):
    for _ in range(x):
        bot.hotkey('shift', 'tab')  

def press_backspace(x):
    for _ in range(x):
        bot.press('backspace')  

def press_right(x):
    for _ in range(x):
        bot.press('right')  

value = df.at[0, 'Norma']
pyperclip.copy(value)

bot.hotkey('ctrl', 'a')
bot.sleep(1)
bot.hotkey('ctrl', 'v')
bot.sleep(1)
bot.typewrite('-032')
bot.sleep(1)
bot.press('enter')
bot.sleep(1)
bot.press('enter')
bot.sleep(1)
bot.press('right')
bot.sleep(1)
press_backspace(3)
bot.sleep(1)
bot.typewrite('32')
bot.sleep(1)
press_tab(10)
bot.sleep(1)
bot.typewrite('17.02.2025')
bot.sleep(1)
press_shtab(7)
bot.sleep(1)
press_right(3)
bot.sleep(1)
bot.press('enter')
bot.sleep(1)
bot.press('tab')
bot.sleep(1)
bot.press('enter')
bot.sleep(1)
bot.typewrite('6854D110-412')
bot.sleep(1)
bot.press('enter')
bot.sleep(1)
bot.press('tab')
bot.sleep(1)
bot.press('enter')
bot.sleep(1.5)
bot.press('enter')
bot.sleep(1.5)
bot.hotkey('ctrl', 's')