# ===== LIBRARIES =====

import pyautogui as bot
import pyperclip as pc
import pandas as pd
import time

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.75

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = 'ApontamentoRodrigo.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
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

def apointment_verifier():
    press_key('tab', 4)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)

    text = pc.paste()
    text = text.strip()

    if text != 'H':
        return False
    else:
        return True

def wcenter_verifier():
    press_key('tab', 2)

    not_empty = True

    while not_empty:
        press_key('ctrla', 1)
        press_key('ctrlc', 1)

        wcenter = pc.paste()
        wcenter = wcenter.strip()

        if wcenter != 'FF78012':
            wcenter = False
        else:
            press_key('tab', 1)
            press_key('ctrla', 1)
            press_key('ctrlc', 1)
            press_key('stab', 1)
            press_key('down', 1)

    press_key('stab', 4)

def open_diagram():
    if wait_event('images/DIAGRAM_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Diagram screen not found <|\n')
    
    lp = df.at[line, 'LPs']
    bot.typewrite(str(lp))
    press_key('f7', 1)

def verify_lp():
    if wait_event('images/DIAGRAM_2.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 2º Diagram screen not found <|\n')

    have_apointment = apointment_verifier()

    if not have_apointment:
        wcenter_verifier()
    else:
        press_key('f3', 1)

def create_apointment():
    bot.typewrite('APS - Rodrigo - 23.07.2026')
    press_key('tab', 2)
    bot.typewrite('H')
    press_key('tab', 2)
    bot.typewrite('FCTMEO')
    press_key('tab', 2)
    bot.typewrite('100')
    press_key('tab', 1)
    bot.typewrite('025PROJ')
    bot.sleep(0.3)

def save_line():
    press_key('ctrls', 1)
    bot.sleep(2.55)
    press_key('tab', 1)
    press_key('enter', 1)
    bot.sleep(1.55)
    bot.typewrite('92866849')
    press_key('enter', 1)
    bot.sleep(1.75)
    press_key('tab', 1)
    bot.sleep(0.75)
    press_key('enter', 1)
    bot.sleep(1.75)

    if wait_event('images/WARNING.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> Diagram screen not found <|\n')

    press_key('enter', 1)
    
    df.at[line, 'Status'] = 'Apointing created!'
    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')         
    bot.sleep(3)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 37
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        open_diagram()
        verify_lp()
        create_apointment()
        save_line()

        line += 1

    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    bot.alert(title='BotText', text='Program successfully completed')