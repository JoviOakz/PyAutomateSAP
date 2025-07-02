# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.25

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

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

    bot.typewrite('6854D110-434')

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

def process_lines():
    global line
    global norm

    for _ in range(repeat_count):
        norm = df.at[line, 'Norma']
        qty = df.at[line, 'Quantidade']

        press_key('tab', 3)
        bot.typewrite(str(norm) + '-001')
        press_key('shtab', 3)

        if norm == 4729106784:
            serie = 42
        else:
            serie = 2

        for __ in range(qty):
            main_function(serie)
            serie += 1
            bot.sleep(2.5)

        line += 1

# ===== PROGRAM CONFIGURATION =====

install_location_qty = 60
line = 0
repeat_count = install_location_qty - line

# ===== MAIN =====

def main():
    process_lines()
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()