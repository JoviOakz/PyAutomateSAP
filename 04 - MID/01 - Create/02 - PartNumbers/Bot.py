# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.15

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def exists_verification():
    bot.click(540, 450)
    bot.sleep(0.3)
    press_key('ctrla', 1)
    bot.typewrite(str(part_number))
    bot.sleep(0.3)
    press_key('enter', 1)

    bot.sleep(1)

    try:
        typeNumber_exist = list(bot.locateAllOnScreen('images/EXIST.png', grayscale=True, confidence=0.9))
        if typeNumber_exist:
            return True
        
    except Exception:
        print('Type number doesn\'t exist!')

# ===== PROGRAM CONFIGURATION =====

part_number_qty = 29
line = 0
repeat_count = part_number_qty - line

# ===== MAIN =====

def main():
    global line
    global part_number

    for _ in range(repeat_count):
        exist = False

        part_number = df.at[line, 'K']  # W -> 764 | X -> 772 | S -> 775 | K -> 29

        exist = exists_verification()

        if not exist:
            bot.click(1870, 1074)
            bot.sleep(0.3)
            bot.click(790, 450)
            bot.sleep(0.3)
            bot.typewrite(str(part_number))
            bot.sleep(0.3)
            press_key('enter', 1)
            bot.sleep(1.25)
            press_key('enter', 1)

        line += 1

    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()