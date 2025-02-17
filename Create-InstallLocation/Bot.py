import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.125

bot.click(1802, 14)

line = 0

excel_path = "LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def main_function(serie):
    bot.hotkey('ctrl', 'a')

    if serie < 10:
        bot.typewrite(str(norm) + '-00' + serie)
    else:
        bot.typewrite(str(norm) + '-0' + serie)
    
    bot.sleep(0.3)
    press_key('enter', 1)
    bot.sleep(0.3)
    press_key('enter', 1)
    bot.sleep(0.5)

    bot.press('right')
    press_key('backspace', 3)

    if serie < 10:
        bot.typewrite('0' + serie)
    else:
        bot.typewrite(serie)

    press_key('tab', 10)
    bot.sleep(0.3)
    bot.typewrite('17.02.2025')
    bot.sleep(0.3)
    press_key('shtab', 7)
    bot.sleep(0.3)
    press_key('right', 3)
    bot.sleep(0.3)
    bot.press('enter')
    bot.sleep(0.5)

    bot.press('tab')
    bot.sleep(0.3)
    bot.press('enter')
    bot.sleep(0.5)
    bot.typewrite('6854D110-412')
    bot.sleep(0.3)
    bot.press('enter')
    bot.sleep(0.5)
    bot.press('tab')
    bot.sleep(0.3)
    bot.press('enter')
    bot.sleep(0.5)
    bot.press('enter')
    bot.sleep(1.25)
    
    bot.hotkey('ctrl', 's')

for _ in range(16):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']

    press_key('tab', 3)

    bot.typewrite(str(norm) + '001')

    press_key('shtab', 3)

    if norm == 4718301303:
        serie = 34
    else:
        serie = 2

    for __ in range(qty):
        main_function(serie)
        serie += 1