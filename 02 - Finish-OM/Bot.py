# ===== README =====
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../../' BEFORE THE PATH OF PDF

# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.75

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../Open-OMs.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlf12':
            bot.hotkey('ctrl', 'f12')
        elif key == 'ctrlshf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        else:
            bot.press(key)

def open_om():
    om_value = df.at[line, 'OM']
    bot.typewrite(str(om_value))
    press_key('enter', 1)
    bot.sleep(3)

def ence_exist():
    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)

        if have_ence:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            bot.sleep(2.25)
            press_key('f3', 1)
            bot.sleep(1.5)
            
            return 1

    except Exception:
        print('❌ Not finished yet!')

def tec_complete():
    press_key('ctrlf12', 1)
    bot.sleep(2)
    press_key('enter', 1)
    bot.sleep(3)
    press_key('enter', 1)
    bot.sleep(3)

def com_complete():
    press_key('ctrlshf12', 1)
    bot.sleep(1.5)

    try:
        conclude = bot.locateOnScreen('images/CONCLUIR.png', grayscale=True, confidence=0.9)
        
        if conclude:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            bot.sleep(2.25)
            press_key('enter', 1)

    except Exception:
        try:
            error = bot.locateOnScreen('images/ERROR.png', grayscale=True, confidence=0.9)

            if error:
                df.at[line, 'Status'] = 'Ordem pendente!'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                bot.sleep(2.25)
                press_key('f12', 1)
                bot.sleep(1)
                press_key('f12', 1)
                bot.sleep(1)
                press_key('f12', 1)

        except Exception:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            print('⚠️ Warning not found!')

    bot.sleep(2.25)

# ===== PROGRAM CONFIGURATION =====

om_qty = 15
line = 0
repeat_qty = om_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        jump_process = 0

        open_om()
        jump_process = ence_exist()

        if jump_process != 1:
            try:
                tec_finished = bot.locateOnScreen('images/BANDEIRA.png', grayscale=True, confidence=0.9)

                if tec_finished:
                    tec_complete()

            except Exception:
                print('Already technically completed!')

            com_complete()

        line += 1

    bot.alert(title='BotText', text='Program terminated!')

if __name__ == '__main__':
    main()