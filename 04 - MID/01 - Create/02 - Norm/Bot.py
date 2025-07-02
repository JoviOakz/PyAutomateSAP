# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.3

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
        else:
            bot.press(key)

def enter_material(norm):
    bot.click(570, 400)
    bot.sleep(0.3)
    bot.typewrite(str(norm))
    bot.sleep(0.3)
    press_key('tab', 1)
    bot.sleep(0.3)
    bot.typewrite(str(norm))
    bot.sleep(0.3)
    press_key('tab', 3)

def series_function(serie, time, value):
    if serie % 2 != 0:
        if serie < 10:
            bot.typewrite('0' + str(serie))
        else:
            bot.typewrite(str(serie))
        
        press_key('tab', 2)
        bot.typewrite(str(time))
        press_key('tab', 1)
        bot.typewrite(str(value))
        press_key('enter', 1)
        bot.sleep(0.15)
    
    else:
        bot.typewrite(str(value))
        press_key('shtab', 1)
        bot.typewrite(str(time))
        press_key('shtab', 2)

        if serie < 10:
            bot.typewrite('0' + str(serie))
        else:
            bot.typewrite(str(serie))

        press_key('enter', 1)
        bot.sleep(0.15)

def save():
    bot.click(1870, 936)

# ===== PROGRAM CONFIGURATION =====

norm_qty = 1
line = 0
repeat_count = norm_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_count):
        norm = df.at[line, 'Norma']
        qty = df.at[line, 'Quantidade']
        time = df.at[line, 'Tempo']
        value = df.at[line, 'Valor']
        serie = 1

        enter_material(norm)

        for __ in range(qty):
            series_function(serie, time, value)
            serie += 1

        line += 1
        save()

if __name__ == '__main__':
    main()