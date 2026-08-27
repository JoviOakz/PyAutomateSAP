# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
import time

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.85

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../99 - Excels/Open-OMs.xlsx'
df = pd.read_excel(
    EXCEL_PATH,
    engine='openpyxl',
    dtype={
        'Status': str
    }
)

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrld':
            bot.hotkey('ctrl', 'down')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrlf12':
            bot.hotkey('ctrl', 'f12')
        elif key == 'ctrlsf12':
            bot.hotkey('ctrl', 'shift', 'f12')
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

def open_om():
    if wait_event('images/ORDER_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Order screen not found <|\n')
    
    bot.typewrite(str(df.at[line, 'OM']))
    press_key('enter', 1)

def verify_status():
    if wait_event('images/ORDER_2.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 2º Order screen not found <|\n')

    bot.click(125, 160)
    press_key('ctrld', 1)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    status = pc.paste().strip()

    if 'LIB' in status:
        return 'LIB'
    elif 'ENTE' in status:
        return 'ENTE'
    elif 'ENCE' in status:
        return 'ENCE'
    else:
        return 'ABER'

def tclose_om():
    press_key('ctrlf12', 1)

    if wait_event('images/ORDER_3.png', timeout=2.5):
        press_key('enter', 1)

        if wait_event('images/WARNING_1.png', timeout=2.5):
            press_key('enter', 1)

        if wait_event('images/WARNING_2.png', timeout=0.85):
            press_key('enter', 1)

        if wait_event('images/ERROR_4.png', timeout=0.85):
            press_key('f3', 1)

            if wait_event('images/ERROR_3.png', timeout=0.85):
                press_key('enter', 1)

            df.at[line, 'Status'] = 'Caso a parte'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            return False

        return True
        
    else:
        if wait_event('images/ERROR_1.png', timeout=1):
            press_key('f3', 1)
            df.at[line, 'Status'] = 'Compromisso pendente'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            return False
        
        else:
            if wait_event('images/ERROR_2.png', timeout=1):
                press_key('f3', 1)
                df.at[line, 'Status'] = 'Encerrar na Z22I150'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                return False
            
            else:
                bot.alert(title='Warning', text='Script error found!')
                df.at[line, 'Status'] = 'Error'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                raise ValueError('\n\n------------- Error: -------------\n|> 1º Close screen not found <|\n')

def close_om():
    press_key('ctrlsf12', 1)

    if wait_event('images/ORDER_4.png', timeout=2):
        press_key('enter', 1)

    if wait_event('images/ORDER_5.png', timeout=2):
        press_key('enter', 1)

# ===== PROGRAM CONFIGURATION =====

om_qty = len(df['OM'])
line = (df['Status'].notna()).sum()
repeat_qty = om_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    if (df['Status'].isna()).any():
        for _ in range(repeat_qty):
            open_om()
            status = verify_status()

            match status:
                case 'ABER':
                    press_key('f3', 1)

                    if wait_event('images/ERROR_3.png'):
                        press_key('enter', 1)

                    df.at[line, 'Status'] = 'Ordem aberta'
                    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

                case 'LIB':
                    can_close = tclose_om()

                    if can_close:
                        if wait_event('images/ORDER_1.png'):
                            pass
                        else:
                            bot.alert(title='Warning', text='Script error found!')
                            df.at[line, 'Status'] = 'Error'
                            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                            raise ValueError('\n\n------------- Error: -------------\n|> 1º Order screen not found <|\n')

                        bot.sleep(0.85)
                        press_key('enter', 1)

                        if wait_event('images/ORDER_2.png'):
                            pass
                        else:
                            bot.alert(title='Warning', text='Script error found!')
                            df.at[line, 'Status'] = 'Error'
                            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                            raise ValueError('\n\n------------- Error: -------------\n|> 2º Order screen not found <|\n')

                        close_om()
                        df.at[line, 'Status'] = 'Encerrado'
                        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

                case 'ENTE':
                    close_om()
                    df.at[line, 'Status'] = 'Encerrado'
                    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

                case 'ENCE':
                    press_key('f3', 1)
                    df.at[line, 'Status'] = 'Encerrado'
                    df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

            line += 1

    bot.alert(title='BotText', text='Programa encerrado!')