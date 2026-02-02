# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.25

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/Cadastro-LPs.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlf12':
            bot.hotkey('ctrl', 'f12')
        elif key == 'ctrlshf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        else:
            bot.press(key)

# ===== PROGRAM CONFIGURATION =====

om_qty = 1
line = 0
repeat_qty = om_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        jump_process = 0

        if jump_process != 1:
            try:
                tec_finished = bot.locateOnScreen('images/BANDEIRA.png', grayscale=True, confidence=0.9)

            except Exception:
                print('Already technically completed!')

        line += 1

if __name__ == '__main__':
    main()
    bot.alert(title='BotText', text='Programa encerrado!')