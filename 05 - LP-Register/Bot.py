# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

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
        else:
            bot.press(key)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 1
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        bot.typewrite('LP-055036')
        press_key('enter', 1)
        bot.sleep(1)

        bot.PAUSE = 0.35
        press_key('tab', 3)
        bot.sleep(1)
        press_key('ctrlf9', 1)
        bot.sleep(1.5)
        press_key('down', 2)
        press_key('tab', 1)
        press_key('down', 2)
        
        bot.PAUSE = 1.5

        press_key('ctrla', 1)
        bot.typewrite('99999999')

        line += 1

if __name__ == '__main__':
    main()
    # bot.alert(title='BotText', text='Programa encerrado!')