# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
from datetime import date
import time

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.75

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/LP-Register.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'ctrle':
            bot.hotkey('ctrl', 'enter')
        elif key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'ctrlstab':
            bot.hotkey('ctrl', 'shift', 'tab')
        else:
            bot.press(key)

def wait_event(img, region=None, timeout=10):
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            local = bot.locateOnScreen(img, region=region, grayscale=True, confidence=0.9)

            if local:
                return local
        except:
            pass

        time.sleep(0.5)
    return None

def open_lp():
    if wait_event('images/PROJECT_BUILDER_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Project Builder screen not found <|\n')
    
    press_key('alt', 1)
    press_key('down', 2)
    press_key('enter', 1)

    if wait_event('images/PROJECT_BUILDER_2.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> Open project window not found <|\n')

    bot.typewrite(str(df.at[line, 'LP']))
    press_key('enter', 1)

def close_project():
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('enter', 1)
    bot.sleep(0.75)

    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('enter', 1)
    bot.sleep(0.75)

    press_key('ctrls', 1)

def close_diagram():
    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 4)
    press_key('right', 1)
    press_key('enter', 1)
    bot.sleep(0.75)

    press_key('alt', 1)
    press_key('right', 1)
    press_key('down', 2)
    press_key('right', 1)
    press_key('down', 6)
    press_key('right', 1)
    press_key('enter', 1)
    bot.sleep(0.75)

    press_key('ctrlstab', 3)
    press_key('up', 2)
    press_key('ctrle', 1)
    bot.sleep(0.75)

def check_status():
    if wait_event('images/PROJECT_BUILDER_3.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 2º Project Builder screen not found <|\n')

    press_key('ctrltab', 4)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    press_key('enter', 1)

    project_status = pc.paste()

    # Inserir o Status -> LBPA

    if 'ABER' in project_status:
        return 'ABER'
    elif 'LIB' in project_status:
        return 'LIB'
    elif 'ENTE' in project_status:
        return 'ENTE'
    else:
        return 'ENCE'

def verify_diagram():
    press_key('ctrlstab', 3)
    press_key('down', 1)
    press_key('right', 1)
    press_key('ctrlstab', 3)
    press_key('up', 1)
    press_key('down', 2)
    press_key('ctrle', 1)
    bot.sleep(0.75)

    press_key('ctrltab', 1)
    press_key('ctrlstab', 1)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)

    lp = pc.paste()

    if '-' in lp:
        print('a')

def process_lp(status):
    match status:
        case 'ABER':
            close_project()

        case 'LIB':
            verify_diagram()

        case 'ENTE':
            print('3')

        case 'ENCE':
            print('4')

# ===== PROGRAM CONFIGURATION =====

lp_qty = 1
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        open_lp()
        status = check_status()
        process_lp(status)

    bot.alert(title='BotText', text='Program successfully completed')