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
        elif key == 'ctrlenter':
            bot.hotkey('ctrl', 'enter')

def lp_verification():
    press_key('left', 1)
    press_key('alt', 1)
    press_key('down', 2)
    press_key('space', 1)
    bot.typewrite(str(df.at[line, 'LP']))
    press_key('enter', 1)

    lp_nexist = bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7)

    if any(lp_nexist):
        press_key('enter', 1)
        press_key('f12', 1)
        df.at[line, 'Status'] = 'LP doesn\'t exist!'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

        return False

    return True

def check_status():
    press_key('ctrltab', 4)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    press_key('enter', 1)

    stats = pc.paste()
    stats = stats.strip()

    if stats and 'ABER' in stats:
        return 'ABER'
    elif stats and 'ENCE' in stats:
        return 'ENCE'
    elif stats and 'ENTE' in stats:
        return 'ENTE'
    elif stats and 'LIB' in stats:
        return 'LIB'
    else:
        return 'ERROR'
    
def open_project():
    for _ in range(2):
        press_key('alt', 1)
        press_key('right', 1)
        press_key('down', 2)
        press_key('right', 1)
        press_key('down', 4)
        press_key('right', 1)
        press_key('down', 1)
        press_key('space', 1)
        press_key('ctrlstab', 3)
        press_key('down', 1)
        press_key('ctrlenter', 1)
    
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('down', 1)
    press_key('space', 1)
    press_key('ctrltab', 2)
    press_key('space', 1)
    bot.sleep(2)
    
def open_diagram():
    press_key('ctrlstab', 3)
    press_key('stab', 2)
    press_key('space', 1)
    press_key('ctrlstab', 3)
    press_key('down', 2)
    press_key('ctrlenter', 1)
    press_key('ctrltab', 2)
    press_key('space', 1)
    bot.sleep(2)

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

def finish_project():
    press_key('ctrlstab', 8)
    press_key('up', 2)
    press_key('ctrlenter', 1)
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('space', 1)
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('space', 1)

def finish_project_aber():
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('space', 1)
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('space', 1)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 0
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        if lp_verification():
            bot.sleep(2.5)

            stats = check_status()

            if stats == 'ABER':
                finish_project_aber()
                df.at[line, 'Status'] = 'LP finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            elif stats == 'LIB':
                open_diagram()
                close_line()
                finish_project()
                df.at[line, 'Status'] = 'LP finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            elif stats == 'ENTE':
                open_project()
                close_line()
                finish_project()
                df.at[line, 'Status'] = 'LP finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            elif stats == 'ENCE':
                press_key('f3')
                bot.sleep(4)
                df.at[line, 'Status'] = 'Already finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            else:
                df.at[line, 'Status'] = 'Error!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

        line += 1

    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()