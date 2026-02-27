# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
from datetime import datetime, date

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.5

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/Cadastro-LPs.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlf9':
            bot.hotkey('ctrl', 'f9')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        elif key == 'sf1':
            bot.hotkey('shift', 'f1')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'ctrlsf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        else:
            bot.press(key)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 2
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    
    press_key('up', 3)
    press_key('stab', 1)
    press_key('right', 3)
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('tab', 1)
    press_key('down', 3)
    bot.sleep(0.65)