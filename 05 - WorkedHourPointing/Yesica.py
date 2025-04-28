# LIBRARIES
import pyautogui as bot
import pandas as pd

# GLOBAL SOFTWARE SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.75

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "ApontamentoYesica.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

# FUNCTION TO PRESS COMMAND X TIMES
def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

# OPEN DIAGRAM
def open_diagram():
    lp = df.at[line, 'LPs']
    bot.typewrite(str(lp))
    press_key('f7', 1)

    bot.sleep(3)

# VERIFIES IF LP HAS ALREADY A FILLED LINE
def verify_lp():
    try:
        not_exist_lp = list(bot.locateAllOnScreen('../images/LPNOTEXIST.png', grayscale=True, confidence=0.9))

        if not_exist_lp:
            df.at[line, 'Status'] = 'WARNING - LP don\'t exist!'
            
            return True

    except Exception:
        try:
            exist_line = list(bot.locateAllOnScreen('../images/NFILLEDLINE.png', grayscale=True, confidence=0.9))

            if not exist_line:
                df.at[line, 'Status'] = 'Line already used!'

                press_key('f3', 2)

                return True
            
        except Exception:
            df.at[line, 'Status'] = 'Line created!'

# INSERT USER INFORMATION
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

# SAVE THE APOINTMENT LINE
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

# EXCEL CONFIG
lp_qty = 35
line = 0

# REPEAT QUANTITY TO PROGRAM RUN
repeat_qty = lp_qty - line

# MAIN PROGRAM
for _ in range(repeat_qty):
    filled_line = False

    open_diagram()
    verify_lp()

    if not filled_line:
        create_apointment()
        save_line()

    line += 1

df.to_excel(excel_path, index=False, engine='openpyxl')