# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 2

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/Scripts.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlf2':
            bot.hotkey('ctrl', 'f2')
        elif key == 'shiftf2':
            bot.hotkey('shift', 'f2')
        elif key == 'ctrlsf4':
            bot.hotkey('ctrl', 'shift', 'f4')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def enter_pn():
    global line

    press_key('ctrla', 1)
    bot.typewrite(str(df.at[line, 'FERRAMENTA']))
    bot.sleep(1.35)
    bot.PAUSE = 0.25
    press_key('tab', 1)
    press_key('down', 4)
    bot.PAUSE = 2
    press_key('ctrla', 1)
    bot.typewrite('WAG20052026')
    press_key('enter', 3)
    bot.sleep(1.65)

def op_upd():
    global line
    temp_line = line
    op_qty = 0
    isEqual = True

    press_key('ctrlf2', 1)
    press_key('shiftf2', 1)
    press_key('tab', 1)
    press_key('enter', 1)

    pn_number = int(df.at[line, 'FERRAMENTA'])

    while isEqual:
        comparative_pn = int(df.at[temp_line, 'FERRAMENTA'])
        temp_line += 1
        op_qty += 1

        if pn_number != comparative_pn:
            isEqual = False

    op_qty -= 1

    for _ in range(op_qty):
        press_key('ctrlsf4', 1)
        bot.sleep(1.35)
        bot.typewrite(str(df.at[line, 'OP']))
        press_key('tab', 1)
        bot.PAUSE = 0.25
        press_key('down', 1)
        bot.PAUSE = 2
        bot.typewrite(str(df.at[line, 'CODE']))
        press_key('enter', 1)
        bot.sleep(1.15)
        df.at[line, 'STATUS'] = 'Cadastrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    
        line += 1

def save_pn():
    press_key('ctrls', 1)

# ===== PROGRAM CONFIGURATION =====

# line_qty = 248
# line = 0
repeat_qty = line_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        enter_pn()
        op_upd()
        save_pn()

    bot.alert(title='BotText', text='Programa encerrado!')