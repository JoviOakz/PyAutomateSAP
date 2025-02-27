import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.3

bot.click(1802, 14)

excel_path = "../Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 0

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def save():
    bot.click(1850, 946)
    bot.sleep(2)

    TELA DE SUCCESS



for _ in range(23):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']
    name = df.at[line, 'Nome']

    bot.click(1850, 946)

    bot.sleep(1.5)

    bot.click(500, 390)
    bot.typewrite(str(name))
    press_key('tab', 2)
    bot.typewrite(str(norm))
    bot.click(500, 520)
    press_key('tab', 1)

    bot.typewrite('SERIE')
    
    press_key('tab', 1)
    bot.typewrite('100000')
    press_key('tab', 1)
    bot.typewrite('15')
    press_key('tab', 1)
    bot.moveTo(500, 580, 0.5)
    bot.scroll(-1000)
    bot.click()
    press_key('tab', 1)
    bot.moveTo(920, 580, 0.5)
    bot.scroll(-6000)
    bot.click()
    press_key('tab', 1)
    bot.typewrite('12')
    bot.click(1272, 724)
    press_key('tab', 2)

    bot.typewrite('SERIE')
    
    bot.click(1436, 1006)
    bot.click(1436, 1006)
    bot.scroll(-6000)
    bot.click(1400, 766)
    