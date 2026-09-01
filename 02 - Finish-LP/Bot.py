# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
from datetime import datetime
import time

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.85

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../99 - Excels/Open-LPs.xlsx'
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
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        elif key == 'sspace':
            bot.hotkey('shift', 'space')
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

def wait_event(img, region=None, timeout=7.5):
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

    press_key('stab', 1)
    press_key('alt', 1)
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

    bot.PAUSE = 0.15
    press_key('ctrltab', 4)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    press_key('enter', 1)
    bot.sleep(0.85)
    bot.PAUSE = 0.85

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

def verify_diagram():
    bot.PAUSE = 0.15
    press_key('ctrlstab', 3)
    press_key('down', 1)
    press_key('right', 1)
    bot.sleep(1.75)
    press_key('ctrle', 1)
    bot.sleep(2)
    press_key('ctrlstab', 3)
    press_key('down', 1)
    press_key('ctrle', 1)
    bot.sleep(2)
    press_key('ctrltab', 1)
    press_key('ctrlstab', 1)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    bot.PAUSE = 0.85

    diagram = pc.paste().strip()

    if '-' not in diagram:
        return True
    else:
        return False

def close_diagram():
    press_key('ctrltab', 2)
    press_key('enter', 1)
    bot.sleep(1.25)
    bot.PAUSE = 0.15
    press_key('tab', 6)

    content = True
    total_lines = 0
    diagram_line = 0
    repeat_process = 0

    while content:
        press_key('ctrla', 1)
        press_key('ctrlc', 1)

        wcenter = pc.paste().strip()

        if wcenter == 'FF78012':
            pass
        elif 'FCT' in wcenter:
            diagram_line += 1
        else:
            break

        press_key('tab', 1)
        press_key('ctrlc', 1)
        press_key('stab', 1)
        press_key('down', 1)
        total_lines += 1

    bot.PAUSE = 0.85
    repeat_process = total_lines - diagram_line

    if repeat_process > 0:
        bot.PAUSE = 0.15
        press_key('up', repeat_process)

        for _ in range(repeat_process):
            press_key('sspace', 1)
            press_key('down', 1)

        press_key('ctrltab', 1)
        press_key('tab', 5)
        press_key('enter', 1)
        bot.sleep(1.25)
        bot.PAUSE = 0.85

        for _ in range(repeat_process):
            press_key('down', 1)
            bot.typewrite('92903610')
            bot.PAUSE = 0.35
            press_key('ctrltab', 1)
            press_key('tab', 1)
            press_key('space', 1)
            press_key('down', 1)
            press_key('space', 1)
            press_key('tab', 1)
            press_key('space', 1)
            bot.PAUSE = 0.15
            press_key('ctrlstab', 2)
            press_key('tab', 3)
            press_key('space', 1)
            bot.PAUSE = 0.85

            if wait_event('images/WARNING_2.png'):
                pass
            else:
                bot.alert(title='Warning', text='Script error found!')
                df.at[line, 'Status'] = 'Error'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                raise ValueError('\n\n------------- Error: -------------\n|> Warning window not found <|\n')

            press_key('tab', 1)
            press_key('enter', 1)

            if wait_event('images/WARNING_3.png'):
                pass
            else:
                bot.alert(title='Warning', text='Script error found!')
                df.at[line, 'Status'] = 'Error'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                raise ValueError('\n\n------------- Error: -------------\n|> Warning window not found <|\n')

            press_key('stab', 1)
            press_key('enter', 1)

            if wait_event('images/WARNING_4.png'):
                pass
            else:
                bot.alert(title='Warning', text='Script error found!')
                df.at[line, 'Status'] = 'Error'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                raise ValueError('\n\n------------- Error: -------------\n|> Warning window not found <|\n')

            press_key('tab', 1)
            press_key('enter', 1)
            bot.sleep(1.25)

    bot.PAUSE = 0.15
    press_key('ctrlstab', 4)
    press_key('enter', 1)
    bot.sleep(1.25)

    press_key('alte', 1)
    press_key('s', 1)
    press_key('t', 1)
    press_key('d', 1)
    bot.sleep(1.25)

    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(1.25)
    bot.PAUSE = 0.85


    if wait_event('images/WARNING_1.png', timeout=2.25):
        press_key('tab', 1)
        press_key('enter', 1)
        bot.sleep(0.85)
        press_key('ctrls', 1)
        df.at[line, 'Status'] = 'Apropriar na CJ88'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        return

    if repeat_process > 0:
        if wait_event('images/WARNING_5.png'):
            press_key('enter', 1)
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> Warning window not found <|\n')

    press_key('ctrlstab', 3)
    press_key('up', 2)
    press_key('ctrle', 1)
    bot.sleep(1.25)

    close_project()

def close_project():
    bot.PAUSE = 0.15
    press_key('alte', 1)
    press_key('s', 1)
    press_key('t', 1)
    press_key('d', 1)
    bot.sleep(1.25)

    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(1.25)
    bot.PAUSE = 0.85

    if wait_event('images/WARNING_1.png', timeout=2.25):
        press_key('tab', 1)
        press_key('enter', 1)
        df.at[line, 'Status'] = 'Apropriar na CJ88'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    else:
        df.at[line, 'Status'] = 'Encerrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

    bot.sleep(0.85)
    press_key('ctrls', 1)

def process_lp(status):
    match status:
        case 'ABER':
            close_project()

        case 'LBPA':
            diagram_exist = verify_diagram()

            if diagram_exist:
                close_diagram()
            else:
                press_key('f3', 1)
                df.at[line, 'Status'] = 'Diagrama não encontrado'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

        case 'LIB':
            diagram_exist = verify_diagram()

            if diagram_exist:
                close_diagram()
            else:
                press_key('f3', 1)
                df.at[line, 'Status'] = 'Diagrama não encontrado'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

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

    if wait_event('images/ACC.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Project Builder screen not found <|\n')

    bot.typewrite('0010')
    press_key('enter', 1)

    if wait_event('images/APPROPRIATION_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')

    bot.typewrite(str(df.at[line, 'LP']))
    press_key('tab', 2)
    press_key('down', 1)
    press_key('space', 1)
    press_key('down', 1)
    press_key('space', 1)
    press_key('tab', 1)
    bot.typewrite(str(int(datetime.now().strftime('%m'))))
    press_key('tab', 1)
    bot.typewrite(str(int(datetime.now().strftime('%m'))))
    press_key('tab', 1)
    bot.typewrite(str(int(datetime.now().strftime('%Y'))))
    press_key('tab', 1)
    bot.typewrite(datetime.today().strftime('%d.%m.%Y'))
    press_key('enter', 1)

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

    if wait_event('images/APPROPRIATE.png', timeout=2.25):
        df.at[line, 'Status'] = 'Apropriado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    else:
        if wait_event('images/APPROPRIATION_2.png'):
            df.at[line, 'Status'] = 'Apropriado'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            press_key('f3')
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º Appropriation screen not found <|\n')

def cj20n_config():
    if wait_event('images/PROJECT_BUILDER_1.png', timeout=2.25):
        pass
    else:
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
    if wait_event('images/PROJECT_BUILDER_3.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 2º Project Builder screen not found <|\n')

    bot.PAUSE = 0.15
    press_key('ctrlstab', 3)
    press_key('down', 1)
    press_key('right', 1)
    bot.sleep(1.75)
    press_key('ctrle', 1)
    bot.sleep(2)
    press_key('ctrlstab', 3)
    press_key('down', 1)
    press_key('ctrle', 1)
    bot.sleep(2)
    bot.PAUSE = 0.15
    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(1.25)
    bot.PAUSE = 0.85

    if wait_event('images/WARNING_1.png', timeout=2.25):
        press_key('tab', 1)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('f3', 1)
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        return

    bot.sleep(0.85)
    bot.PAUSE = 0.15
    press_key('ctrlstab', 3)
    press_key('up', 2)
    press_key('ctrle', 1)
    bot.sleep(2)
    press_key('alte', 1)
    press_key('s', 1)
    press_key('t', 1)
    press_key('d', 1)
    bot.sleep(1.25)
    press_key('alte', 1)
    press_key('s', 1)
    press_key('a', 1)
    press_key('d', 1)
    bot.sleep(1.25)
    bot.PAUSE = 0.85

    if wait_event('images/WARNING_1.png', timeout=2.25):
        press_key('tab', 1)
        press_key('enter', 1)
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
    else:
        df.at[line, 'Status'] = 'Encerrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

    bot.sleep(0.85)
    press_key('ctrls', 1)

# ===== PROGRAM CONFIGURATION =====

lp_qty = len(df['LP'])
line = (df['Status'].notna()).sum()
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    if (df['Status'].isna()).any():
        for _ in range(repeat_qty):
            open_lp()
            status = check_status()
            process_lp(status)
            line += 1

    if (df['Status'] == 'Apropriar na CJ88').any():
        line = (df['Status'] != 'Apropriar na CJ88').sum()
        repeat_qty = lp_qty - line
        cj88_config()

        for _ in range(repeat_qty):
            lp_status = str(df.at[line, 'Status'])

            if lp_status == 'Apropriar na CJ88':
                appropriate_lp()

            line += 1

    if (df['Status'] == 'Apropriado').any():
        line = (df['Status'] != 'Apropriado').sum()
        repeat_qty = lp_qty - line
        cj20n_config()

        for _ in range(repeat_qty):
            lp_status = str(df.at[line, 'Status'])

            if lp_status == 'Apropriado':
                open_lp()
                ence_project()

            line += 1

    bot.alert(title='BotText', text='Program successfully completed')