# LIBRARIES
import pyautogui as bot
import pandas as pd

# GLOBAL SOFTWARE SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.5

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "ApontamentoDiego.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

# FUNCTION TO PRESS COMMAND X TIMES
def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

#
def open_diagram():
    lp = df.at[line, 'LPs']
    bot.typewrite(str(lp))
    press_key('f7', 1)

#
def create_apointment():
    press_key('tab', 2)
    bot.typewrite('Projetista Diego - 28.03.2025')
    press_key('tab', 2)
    bot.typewrite('H')
    press_key('tab', 2)
    bot.typewrite('FCTVID')
    press_key('tab', 2)
    bot.typewrite('100')
    press_key('tab', 1)
    bot.typewrite('025PROJ')

    bot.sleep(0.3)

#
def save_line():
    press_key('ctrls', 1)
    bot.sleep(2)
    press_key('tab', 1)
    press_key('enter', 1)
    bot.sleep(1)
    bot.typewrite('92903130')
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('tab', 1)
    bot.sleep(0.3)
    press_key('enter', 1)
    
    bot.sleep(2)

# EXCEL CONFIG
lp_qty = 20
line = 0

# REPEAT QUANTITY TO PROGRAM RUN
repeat_qty = lp_qty - line

# MAIN PROGRAM
for _ in range(repeat_qty):
    open_diagram()
    create_apointment()
    save_line()

    line += 1