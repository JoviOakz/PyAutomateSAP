import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.35

bot.click(1802, 14)

line = 20

excel_path = "ApontamentoYesica.xlsm"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

for _ in range(80):
    lp = df.at[line, 'Elemento PEP']

    bot.typewrite(str(lp))

    press_key('f7', 1)

    bot.sleep(3)

    press_key('tab', 2)
    bot.typewrite('Planejadora Yesica - 20.02.2025')
    press_key('tab', 2)
    bot.typewrite('H')
    press_key('tab', 2)
    bot.typewrite('FCTLIY')
    press_key('tab', 2)
    bot.typewrite('100')
    press_key('tab', 1)
    bot.typewrite('025PROJ')
    press_key('ctrls', 1)

    bot.sleep(2)

    press_key('tab', 1)
    press_key('enter', 1)

    bot.sleep(0.75)
    
    bot.typewrite('92886895')
    press_key('enter', 1)

    bot.sleep(1.25)
    
    press_key('tab', 1)

    bot.sleep(0.3)
    
    press_key('enter', 1)
    
    bot.sleep(1)

    line += 1