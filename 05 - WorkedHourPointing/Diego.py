import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.5

bot.click(1802, 14)

excel_path = "ApontamentoDiego.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

qty = 37
line = 0

repeat_count = qty - line

for _ in range(repeat_count):
    lp = df.at[line, 'LPs']

    bot.typewrite(str(lp))

    press_key('f7', 1)

    bot.sleep(3)

    press_key('tab', 2)
    bot.typewrite('Planejadora Yesica - 25.03.2025')
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

    bot.sleep(1)
    
    bot.typewrite('92886895')
    press_key('enter', 1)

    bot.sleep(1.25)
    
    press_key('tab', 1)

    bot.sleep(0.3)
    
    press_key('enter', 1)
    
    bot.sleep(2)

    line += 1