# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.75

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = 'ApontamentoYesica.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def open_diagram():
    lp = df.at[line, 'LPs']
    bot.typewrite(str(lp))
    press_key('f7', 1)
    bot.sleep(3)

def verify_lp():
    try:
        not_exist_lp = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.9))

        if not_exist_lp:
            df.at[line, 'Status'] = 'LP doesn\'t exist!'
            
            return True

    except Exception:
        try:
            exist_line = list(bot.locateAllOnScreen('images/NFILLEDLINE.png', grayscale=True, confidence=0.9))

            if exist_line:
                df.at[line, 'Status'] = 'Line created!'
            
        except Exception:
            df.at[line, 'Status'] = 'Line already used!'
            press_key('f3', 2)

            return True

def create_apointment():
    press_key('tab', 2)
    bot.typewrite('Planejadora Yesica - 28.04.2025')
    press_key('tab', 2)
    bot.typewrite('H')
    press_key('tab', 2)
    bot.typewrite('FCTLIY')
    press_key('tab', 2)
    bot.typewrite('100')
    press_key('tab', 1)
    bot.typewrite('025PROJ')
    bot.sleep(0.3)

def save_line():
    press_key('ctrls', 1)
    bot.sleep(2)
    press_key('tab', 1)
    press_key('enter', 1)
    bot.sleep(1)
    bot.typewrite('92886895')
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('tab', 1)
    bot.sleep(0.3)
    press_key('enter', 1)
    bot.sleep(3)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 35
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
            create_apointment()
            save_line()

        line += 1

    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    bot.alert(title='BotText', text='Program terminated!')

if __name__ == '__main__':
    main()