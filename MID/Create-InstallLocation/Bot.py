import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.click(1802, 14)

excel_path = "LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        else:
            bot.press(key)

def main_function(serie):
    press_key('ctrla', 1)

    if serie < 10:
        bot.typewrite(str(norm) + '-00' + str(serie))
    else:
        bot.typewrite(str(norm) + '-0' + str(serie))
    
    bot.sleep(0.5)
    press_key('enter', 1)
    bot.sleep(2)
    press_key('enter', 1)
    bot.sleep(2)

    press_key('right', 1)
    press_key('backspace', 3)

    if serie < 10:
        bot.typewrite('0' + str(serie))
    else:
        bot.typewrite(str(serie))

    press_key('tab', 10)
    bot.sleep(0.5)

    bot.typewrite('10.03.2025')

    bot.sleep(0.5)
    press_key('shtab', 7)
    bot.sleep(0.5)
    press_key('right', 3)
    bot.sleep(0.5)
    press_key('enter', 1)
    bot.sleep(2)

    press_key('tab', 1)
    bot.sleep(0.5)
    press_key('enter', 1)
    bot.sleep(1.5)

    if norm == 4729108508 or 4328700313:
        bot.typewrite('6854D120-438')
    else:
        bot.typewrite('6854D120-434')

    bot.sleep(0.5)
    press_key('enter', 1)
    bot.sleep(1.5)
    press_key('tab', 1)
    bot.sleep(0.5)
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('enter', 1)
    bot.sleep(1.25)
    
    bot.hotkey('ctrl', 's')

line = 1

for _ in range(3):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']

    press_key('tab', 3)

    bot.typewrite(str(norm) + '-001')

    press_key('shtab', 3)

    if norm == 4328700313:
        serie = 4
    else:
        serie = 2

    for __ in range(qty):
        main_function(serie)
        serie += 1
        bot.sleep(2.5)

    line += 1