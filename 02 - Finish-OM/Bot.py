# --- README ---
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../../' BEFORE THE PATH OF PDF

# LIBRARIES
import pyautogui as bot
import pandas as pd
import pyperclip

# SOFTWARE GLOBAL SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.6

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "../Open-OMs.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')
    
# FUNCTION TO PRESS COMMAND X TIMES
def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl,' 'v')
        elif key == 'ctrlf12':
            bot.hotkey('ctrl', 'f12')
        elif key == 'ctrlshf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        else:
            bot.press(key)

# OPEN ORDER
def open_om():
    om_value = df.at[line, 'OM']
    bot.typewrite(str(om_value))
    press_key('enter', 1)

    bot.sleep(3)

# TECHNICALLY COMPLETE
def tec_complete():
    press_key('ctrlf12', 1)
    bot.sleep(1.5)
    press_key('enter', 1)
    bot.sleep(2.75)

    press_key('enter', 1)
    bot.sleep(3)

# COMERCIAL COMPLETE
def com_complete():
    press_key('ctrlshf12', 1)

    bot.sleep(1.5)

    try:
        conclude = bot.locateOnScreen('images/CONCLUIR.png', grayscale=True, confidence=0.9)
        
        if conclude:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            bot.sleep(2.25)

            press_key('enter', 1)

    except Exception:
        print('Warning not found!')

    try:
        error = bot.locateOnScreen('images/ERROR.png', grayscale=True, confidence=0.9)
        
        if error:
            df.at[line, 'Status'] = 'Ordem pendente!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            bot.sleep(2.25)

            press_key('f12', 1)
            bot.sleep(1)
            press_key('f12', 1)
            bot.sleep(1)
            press_key('f12', 1)

    except Exception:
        print('Warning not found!')

    bot.sleep(1.5)

# EXCEL CONFIG
lp_qty = 20
line = 0

# REPEAT QUANTITY TO PROGRAM RUN
repeat_qty = lp_qty - line

# MAIN PROGRAM
for _ in range(repeat_qty):
    open_om()

    try:
        tec_finished = bot.locateOnScreen('images/BANDEIRA.png', grayscale=True, confidence=0.9)
        
        if tec_finished:
            tec_complete()

    except Exception:
        print('Already tecnically completed!')

    com_complete()
            
    line += 1