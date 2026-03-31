# ===== LIBRARIES =====

import pyautogui as bot
import pyperclip as pc
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.25

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../../01 - Excels/Open-LPs.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'ctrlstab':
            bot.hotkey('ctrl', 'shift', 'tab')
        elif key == 'sspace':
            bot.hotkey('shift', 'space')
        elif key == 'ctrlr':
            bot.hotkey('ctrl', 'right')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')

def open_project():
    press_key('left', 1)
    press_key('alt', 1)
    press_key('down', 2)
    press_key('space', 1)
    bot.typewrite(str(df.at[line, 'LP']))
    press_key('enter', 1)
    bot.sleep(2.5)

def lp_verification():
    print('teste')

def text_verifier():
    text = pc.paste()
    text = text.strip()

    if text != 'x':
        if len(text) == 7:
            return False, 1
        else:
            return False, 0
    else:
        return True, 0

def close_line():
    empty = False
    fct_counter = 0
    line_counter = 0

    press_key('tab', 6)

    while not empty:
        press_key('ctrlr', 1)
        press_key('backspace', 1)
        bot.typewrite('x')
        press_key('ctrla', 1)
        press_key('ctrlc', 1)
        press_key('backspace', 1)
        press_key('tab', 1)
        press_key('stab', 1)
        press_key('down', 1)

        empty, fct_counter = text_verifier()
        line_counter += 1
    
    press_key('up', line_counter)

    if fct_counter >= 1:
        press_key('down', 1)

    for _ in range(line_counter - fct_counter - 1):
        press_key('sspace', 1)
        press_key('down', 1)

    press_key('ctrltab', 1)
    press_key('tab', 5)
    press_key('space', 1)
    press_key('down', 1)
    bot.typewrite('92903610')
    press_key('down', 2)
    press_key('tab', 1)
    press_key('down', 1)
    press_key('space', 1)
    press_key('down', 1)
    press_key('space', 1)
    press_key('tab', 1)
    press_key('space', 1)
    press_key('enter', 1)
    press_key('tab', 1)
    press_key('space', 1)
    press_key('ctrlstab', 1)
    press_key('space', 1)
    press_key('tab', 1)
    press_key('space', 1)
    press_key('ctrlstab', 4)
    press_key('space', 1)
    bot.sleep(2)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 0
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        open_project()
        if lp_verification():
            print('teste')
    #         if diagram_adjustment():
    #             close_lp()
    #         else:
    #             cancel()
    #     else:
    #         cancel()

        line += 1

    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()