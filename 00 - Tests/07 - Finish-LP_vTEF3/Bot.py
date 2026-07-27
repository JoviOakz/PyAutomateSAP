# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
from datetime import date
import time

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.85

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../../01 - Excels/TEF3.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        if key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'ctrle':
            bot.hotkey('ctrl', 'enter')
        elif key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'ctrlstab':
            bot.hotkey('ctrl', 'shift', 'tab')
        elif key == 'alte':
            bot.hotkey('alt', 'e')
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

    bot.click(1200, 300)
    press_key('alt', 2)
    press_key('b', 1)

    if wait_event('images/PROJECT_BUILDER_2.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> Open project window not found <|\n')

    bot.typewrite(str(df.at[line, 'LP']))
    press_key('enter', 1)

def check_status():
    if wait_event('images/PROJECT_BUILDER_3.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 2º Project Builder screen not found <|\n')

    bot.PAUSE = 0.25
    press_key('ctrltab', 4)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    press_key('enter', 1)
    bot.PAUSE = 0.85
    bot.sleep(0.85)

    project_status = pc.paste()

    if 'ABER' in project_status:
        return 'ABER'
    elif 'LBPA' in project_status:
        return 'LBPA'
    elif 'LIB' in project_status:
        return 'LIB'
    elif 'ENTE' in project_status:
        return 'ENTE'
    else:
        return 'ENCE'

def close_project():
    press_key('alte', 1)
    press_key('s', 1)
    press_key('t', 1)
    press_key('d', 1)
    bot.sleep(0.85)

    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(0.85)

    if wait_event('images/WARNING_1.png', timeout=2.25):
        press_key('tab', 1)
        press_key('enter', 1)
        df.at[line, 'Status'] = 'Apropriar na CJ88'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    else:
        df.at[line, 'Status'] = 'Encerrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

    press_key('ctrls', 1)

def process_lp(status):
    match status:
        case 'ABER':
            close_project()

        case 'LBPA':
            close_project()

        case 'LIB':
            close_project()

        case 'ENTE':
            press_key('f3', 1)
            df.at[line, 'Status'] = 'Apropriar na CJ88'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

        case 'ENCE':
            press_key('f3', 1)
            df.at[line, 'Status'] = 'Encerrado'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

def cj88_config():
    if wait_event('images/PROJECT_BUILDER_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Project Builder screen not found <|\n')
    
    bot.typewrite('/NCJ88')
    press_key('enter', 1)

    if wait_event('images/APPROPRIATION_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')
    
    press_key('tab', 2)
    press_key('down', 1)
    press_key('space', 1)
    press_key('down', 1)
    press_key('space', 1)
    press_key('tab', 1)
    bot.typewrite('7')
    press_key('tab', 1)
    bot.typewrite('7')
    press_key('tab', 1)
    bot.typewrite('2026')
    press_key('tab', 1)
    bot.typewrite(date.today().strftime('%d.%m.%Y'))

def appropriate_lp():
    if wait_event('images/APPROPRIATION_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')
    
    bot.typewrite(str(df.at[line, 'LP']))
    press_key('ctrltab', 2)
    press_key('space', 1)
    press_key('f8', 1)
    df.at[line, 'Status'] = 'Apropriado'
    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

    if wait_event('images/APPROPRIATION_2.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')

    press_key('f3')

def cj20n_config():
    if wait_event('images/APPROPRIATION_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')

    press_key('ctrlstab', 1)
    press_key('tab', 1)
    bot.typewrite('/NCJ20N')
    press_key('enter', 1)

def ence_project():
    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(0.85)

    press_key('ctrls', 1)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 166
line = 115
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        open_lp()
        status = check_status()
        process_lp(status)
        line += 1

    line = 0
    cj88_config()

    for _ in range(repeat_qty):
        lp_status = str(df.at[line, 'Status'])

        if lp_status == 'Apropriar na CJ88':
            appropriate_lp()

        line += 1

    line = 0
    cj20n_config()

    for _ in range(repeat_qty):
        lp_status = str(df.at[line, 'Status'])

        if lp_status == 'Apropriado':
            open_lp()
            ence_project()

        line += 1

    bot.alert(title='BotText', text='Program successfully completed')