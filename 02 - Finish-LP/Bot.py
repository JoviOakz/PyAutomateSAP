# ===== LIBRARIES =====

import pyautogui as bot
import pyperclip as pc
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.65

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/Open-LPs.xlsx'
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
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def lp_verification():
    press_key('left', 1)
    press_key('alt', 1)
    press_key('down', 2)
    press_key('space', 1)
    bot.sleep(1)
    bot.typewrite(str(df.at[line, 'LP']))
    press_key('enter', 1)
    bot.sleep(1.5)

    lp_nexist = None

    try:
        lp_nexist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.9))
    except Exception:
        lp_nexist = []

    if any(lp_nexist):
        press_key('enter', 1)
        bot.sleep(1)
        press_key('f12', 1)
        bot.sleep(1)
        df.at[line, 'Status'] = 'LP doesn\'t exist!'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        return False
    
    else:
        return True

def check_status():
    bot.PAUSE = 0.2

    press_key('ctrltab', 4)
    
    bot.sleep(0.75)
    
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    press_key('enter', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(1.5)

    stats = pc.paste()
    stats = stats.strip()

    try:
        diagram_exist = list(bot.locateAllOnScreen('images/DIAGRAM.png', grayscale=True, region=(30, 216, 54, 244), confidence=0.9))
    except Exception:
        diagram_exist = []

    if any(diagram_exist):
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
    else:
        if stats and 'ENCE' in stats:
            return 'ENCE'
        else:
            return 'ABER'
    
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
        bot.sleep(1.5)
        press_key('ctrlstab', 3)
        press_key('down', 1)
        press_key('ctrlenter', 1)
        bot.sleep(1.5)
    
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('down', 1)
    press_key('space', 1)
    bot.sleep(1.5)
    press_key('ctrltab', 2)
    press_key('space', 1)
    bot.sleep(2)
    
def open_diagram():
    bot.PAUSE = 0.2

    press_key('ctrlstab', 4)
    press_key('end', 1)
    press_key('left', 1)
    press_key('space', 1)
    bot.sleep(2.5)
    press_key('ctrlstab', 3)
    press_key('down', 2)
    press_key('ctrlenter', 1)
    bot.sleep(2)
    press_key('ctrltab', 2)
    press_key('space', 1)
    bot.sleep(2.5)

    bot.PAUSE = 0.65

def text_verifier(fct_counter):
    text = pc.paste()
    text = text.strip()

    if text != 'x':
        if text[:3] == 'FCT':
            return False, 1
        else:
            return False, fct_counter
    else:
        return True, fct_counter
    
def insert_real_data():
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
    bot.sleep(1.5)
    press_key('tab', 1)
    press_key('space', 1)
    bot.sleep(1.5)
    press_key('ctrlstab', 1)

def close_line():
    empty = False
    fct_counter = 0
    line_counter = 0

    bot.PAUSE = 0.2

    press_key('tab', 6)
    bot.sleep(1.5)

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

        empty, fct_counter = text_verifier(fct_counter)
        line_counter += 1

    bot.PAUSE = 0.65

    press_key('up', line_counter)
    line_counter = line_counter - fct_counter - 1

    if line_counter != 0:
        if fct_counter > 0:
            press_key('down', 1)

        for _ in range(line_counter):
            press_key('sspace', 1)
            press_key('down', 1)

        bot.sleep(1)
        press_key('ctrltab', 1)

        bot.PAUSE = 0.2

        press_key('tab', 5)
        
        bot.PAUSE = 0.65

        press_key('space', 1)
        bot.sleep(1.75)

        if line_counter == 1:
            insert_real_data()
            press_key('space', 1)
            bot.sleep(1.5)
            press_key('tab', 1)
            press_key('space', 1)
            bot.sleep(3)
            return True
        else:
            for _ in range(line_counter):
                insert_real_data()
                press_key('tab', 3)
                press_key('space', 1)
                bot.sleep(1.5)
                press_key('tab', 1)
                press_key('space', 1)
                bot.sleep(1.5)
            
            press_key('stab', 1)
            press_key('space', 1)
            bot.sleep(1.5)
            press_key('tab', 1)
            press_key('space', 1)
            bot.sleep(3)
            return True
    else:
        return False
    
def finish_diagram():
    bot.PAUSE = 0.2

    press_key('ctrlstab', 4)
    
    bot.PAUSE = 0.65

    press_key('space', 1)
    bot.sleep(2)
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)
    press_key('ctrltab', 5)
    bot.sleep(2)

def finish_project():
    bot.PAUSE = 0.2

    press_key('ctrlstab', 8)
    
    bot.PAUSE = 0.65

    press_key('up', 2)
    press_key('ctrlenter', 1)
    bot.sleep(2)
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)

def finish_project_aber():
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)
    press_key('alt', 1)
    press_key('right', 1)

    bot.PAUSE = 0.2

    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('space', 1)
    
    bot.PAUSE = 0.65
    
    bot.sleep(2)
    press_key('ctrls', 1)
    bot.sleep(6.5)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 366
line = 80
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
                has_purchase = close_line()
                if has_purchase:
                    finish_project()
                    press_key('space', 1)
                    bot.sleep(2)
                    press_key('ctrls', 1)
                    bot.sleep(6.5)
                else:
                    finish_diagram()
                    finish_project()
                    press_key('ctrls', 1)
                    bot.sleep(6.5)
                df.at[line, 'Status'] = 'LP finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            elif stats == 'ENTE':
                open_project()
                has_purchase = close_line()
                if has_purchase:
                    finish_project()
                    press_key('space', 1)
                    bot.sleep(2)
                    press_key('ctrls', 1)
                    bot.sleep(6.5)
                else:
                    finish_diagram()
                    finish_project()
                    press_key('ctrls', 1)
                    bot.sleep(6.5)
                df.at[line, 'Status'] = 'LP finished!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            elif stats == 'ENCE':
                press_key('f3', 1)
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