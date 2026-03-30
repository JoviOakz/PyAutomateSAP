# ===== LIBRARIES =====

import pyautogui as bot
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

def open_diagram():
    lp = df.at[line, 'LPs']
    bot.typewrite(str(lp))
    press_key('f7', 1)
    bot.sleep(3)

def text_verifier():
    text = pc.paste()
    text = text.strip()

    if text != 'x':
        return True
    else:
        return False

def verify_lp():
    try:
        not_exist_lp = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.9))

        if not_exist_lp:
            df.at[line, 'Status'] = 'LP doesn\'t exist!'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            bot.sleep(2)
            
            return True

    except Exception:
        try:
            h_exist = list(bot.locateAllOnScreen('images/H.png', grayscale=True, confidence=0.9))
            
            if h_exist:
                df.at[line, 'Status'] = 'LP already have pointing!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                bot.sleep(2)

                press_key('f3', 2)

                return True
            
        except Exception:
            bot.PAUSE = 0.1
            press_key('tab', 6)
            bot.PAUSE = 0.65

            result = True

            while result:
                press_key('ctrlr', 1)
                press_key('backspace', 1)
                bot.typewrite('x')
                press_key('ctrla', 1)
                press_key('ctrlc', 1)
                press_key('backspace', 1)
                press_key('tab', 1)
                press_key('stab', 1)
                press_key('down', 1)

                result = text_verifier()

            press_key('up', 1)
            bot.PAUSE = 0.1
            press_key('stab', 6)
            bot.PAUSE = 1.25

            return False

# ===== PROGRAM CONFIGURATION =====

lp_qty = 0
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        filled_line = False

        open_diagram()
        filled_line = verify_lp()

        if not filled_line:
            # create_apointment()
            # save_line()
            print('a')

        line += 1

    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()