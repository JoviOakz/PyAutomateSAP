# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
from datetime import date
import time
import re

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.25

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/LP-Register.xlsx'
# EXCEL_PATH = 'Teste.xlsx'
df = pd.read_excel(
    EXCEL_PATH,
    engine='openpyxl',
    dtype={
        'Status CJ02': str,
        'Status CN21': str
    }
)

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlf9':
            bot.hotkey('ctrl', 'f9')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        elif key == 'sf1':
            bot.hotkey('shift', 'f1')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'winr':
            bot.hotkey('win', 'r')
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

def wbs_element_creation():
    global line

    for _ in range(repeat_qty):
        if wait_event('images/PROJECT_1.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º Project screen not found <|\n')
        
        press_key('ctrla', 1)
        bot.typewrite(df.at[line, 'Elemento PEP'])
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.5
        
        if wait_event('images/PROJECT_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º Project screen not found <|\n')

        press_key('ctrla', 1)
        part_number = re.sub(r'[-./POSpos& ]', '', str(df.at[line, 'Part Number'])).strip()
        pc.copy(str(df.at[line, 'Denominação']))

        if part_number.isdigit():
            bot.typewrite(str(df.at[line, 'Part Number']) + ' - ')

        press_key('ctrlv', 1)

        bot.sleep(1.25)
        press_key('tab', 3)
        bot.sleep(0.85)

        if part_number.isdigit():
            bot.typewrite(str(df.at[line, 'Part Number']) + ' - ')

        press_key('ctrlv', 1)

        bot.sleep(1.5)
        press_key('ctrlf9', 1)

        if wait_event('images/WBS_1.png', region=(470, 470, 180, 50)):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º WBS screen not found <|\n')
        
        press_key('down', 2)
        press_key('tab', 1)
        press_key('down', 2)
        bot.sleep(0.75)
        press_key('ctrla', 1)

        keys = {
            'TEF': '68540012',
            'QMM': '68540007',
            'MFW1': '68540001',
            'MFE2': '68540002',
            'MFE3': '68540003',
        }

        department = df.at[line, 'Departamento Emissor']

        iss_dept = None

        for key, value in keys.items():
            if key in department:
                iss_dept = value
                break
            else:
                iss_dept = '68540028'

        bot.typewrite(iss_dept)
        bot.sleep(1)

        bot.PAUSE = 0.75

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 4)
        press_key('enter', 1)

        if wait_event('images/WBS_2.png', region=(320, 500, 220, 50)):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º WBS screen not found <|\n')

        press_key('tab', 1)
        press_key('down', 3)
        bot.sleep(0.65)

        bot.PAUSE = 1.15

        bot.typewrite(df.at[line, 'Entregar para'])
        bot.sleep(0.85)
        press_key('tab', 2)
        bot.typewrite(str(df.at[line, 'Quantidade']).split('.')[0])
        press_key('tab', 1)
        bot.typewrite('PC')
        press_key('tab', 1)
        press_key('down', 1)
        bot.typewrite(str(df.at[line, 'Custo estimado']).split('.')[0])
        press_key('tab', 1)
        bot.typewrite('BRL')
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.75

        press_key('stab', 4)
        press_key('enter', 1)

        bot.PAUSE = 1.5

        if wait_event('images/PARAMETERS_1.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º Parameters screen not found <|\n')

        press_key('tab', 1)

        liquidation_object = str(df.at[line, 'Objeto de Liquidação']).strip().split('.')[0]
        alocation = str(df.at[line, 'Esquema de Alocação']).strip()

        if len(alocation) == 1:
            alocation = '0' + str(df.at[line, 'Esquema de Alocação']).strip()
        
        if liquidation_object.startswith('685') and len(liquidation_object) == 6:
            bot.typewrite('ZPS001')
        elif liquidation_object.startswith('LP-'):
            bot.typewrite('ZPS007')
            alocation = '07'
        elif liquidation_object.startswith('BM'):
            bot.typewrite('ZPS007')
            alocation = '07'
        else:
            bot.typewrite('ZPS003')
            alocation = '07'

        bot.sleep(1)
        press_key('tab', 1)
        bot.typewrite(alocation)
        bot.sleep(1)

        press_key('f3', 1)

        if wait_event('images/PARAMETERS_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º Parameters screen not found <|\n')
        
        press_key('tab', 1)
        bot.typewrite(str(df.at[line, 'Objeto de Liquidação']).split('.')[0])
        press_key('f3', 1)

        if wait_event('images/WBS_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º Return error <|\n')
        
        press_key('f3', 1)

        if wait_event('images/RETURN.png', region=(180, 190, 130, 50)):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º Return error <|\n')
        
        press_key('sf1', 1)

        if wait_event('images/PROJECT_3.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CJ02'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 3º WBS screen not found <|\n')

        bot.PAUSE = 0.5

        press_key('tab', 1)
        press_key('down', 1)
        press_key('tab', 2)
        bot.typewrite(iss_dept)
        bot.sleep(1)
        press_key('down', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        bot.sleep(1)
        press_key('down', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        bot.sleep(1)
        press_key('alt', 1)
        press_key('right', 1)
        press_key('down', 3)
        press_key('right', 1)
        press_key('enter', 1)
        bot.sleep(1.75)
        press_key('ctrls', 1)
        df.at[line, 'Status CJ02'] = 'Cadastrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        
        line += 1

    bot.sleep(2)

def enter_cn21():
    if wait_event('images/PROJECT_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status CN21'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> Project screen not found <|\n')
    
    press_key('stab', 1)
    press_key('left', 7)

    bot.PAUSE = 0.75

    bot.typewrite('/ncn21')
    press_key('enter', 1)

    if wait_event('images/DIAGRAM_1.png'):
        pass
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status CN21'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> 1º Diagram screen not found <|\n')
    
    press_key('right', 1)

    bot.PAUSE = 0.15

    press_key('left', 3)
    press_key('right', 2)
    
    bot.PAUSE = 0.75

    press_key('tab', 1)
    bot.typewrite('BP01')
    press_key('tab', 1)
    bot.typewrite('6854')
    bot.sleep(1.15)
    press_key('stab', 2)

def mrp_config(line):
    mrp = str(df.at[line, 'Responsável']).strip()
    resp_change = False

    if line > 0:
        previous_mrp = str(df.at[line - 1, 'Responsável']).strip()

        if mrp != previous_mrp:
            resp_change = True

    if line == 0 or resp_change:
        press_key('tab', 3)
        bot.sleep(1.25)

        if mrp == 'Yesica Gonzalez':
            bot.typewrite('I33')
        elif mrp == 'Rodrigo Melo':
            bot.typewrite('I49')
        elif mrp == 'Marcelo Simoes':
            bot.typewrite('I55')
        elif mrp == 'Edson Bento':
            bot.typewrite('I38')
        elif mrp == 'Thais Fischer':
            bot.typewrite('I55')
        elif mrp == 'Joao Franca':
            bot.typewrite('I31')
        else:
            bot.typewrite('I39')

        bot.sleep(1.25)

def diagram_creation():
    global line

    for _ in range(repeat_qty):
        bot.PAUSE = 0.35

        mrp_config(line)
        
        bot.PAUSE = 1.5
        
        press_key('enter', 1)

        if wait_event('images/VALUE.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CN21'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> Value box not found <|\n')

        bot.typewrite(df.at[line, 'Elemento PEP'].replace('-', ''))
        press_key('enter', 1)
        bot.PAUSE = 1.25

        if wait_event('images/DIAGRAM_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CN21'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º Diagram screen not found <|\n')

        press_key('tab', 2)
        press_key('right', 1)
        press_key('enter', 1)

        if wait_event('images/ATTRIBUITION_1.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CN21'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º Attribuition screen not found <|\n')

        press_key('tab', 2)
        bot.typewrite(df.at[line, 'Elemento PEP'].replace('-', ''))
        press_key('enter', 1)
        bot.sleep(1.5)
        press_key('tab', 1)
        press_key('enter', 1)
        bot.sleep(1.5)

        bot.PAUSE = 1.25

        press_key('tab', 2)
        press_key('right', 3)
        press_key('enter', 1)

        if wait_event('images/ATTRIBUITION_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status CN21'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º Attribuition screen not found <|\n')

        press_key('tab', 5)

        bot.PAUSE = 1.35
        
        press_key('ctrla', 1)

        heijunka = ''
        description = ''
        responsible = ''

        part_number = re.sub(r'[-./POSpos& ]', '', str(df.at[line, 'Part Number'])).strip()

        if str(df.at[line, 'Responsável']) == 'Yesica Gonzalez' or str(df.at[line, 'Responsável']) == 'Rodrigo Melo':
            heijunka = 'HEIJUNKA\n'

        if part_number.isdigit():
            description = str(df.at[line, 'Part Number']) + ' - '

        description += str(df.at[line, 'Denominação'])
        responsible = '\nResp. ' + str(df.at[line, 'Responsável'])
        full_text = heijunka + description + responsible

        bot.typewrite(full_text)
        bot.sleep(1.15)
        press_key('ctrlsf12', 1)
        bot.sleep(3)
        press_key('ctrls', 1)
        df.at[line, 'Status CN21'] = 'Cadastrado'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        bot.sleep(3)

        line += 1

# ===== PROGRAM CONFIGURATION =====

lp_qty = len(df['Responsável'])
line = (df['Status CJ02'].notna()).sum()
repeat_qty = lp_qty - line

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# press_key('winr', 1)
# bot.typewrite('saplogon')
# press_key('enter', 1)

# # verificação saplogon
# if wait_event('images/ATTRIBUITION_2.png'):
#     pass
# else:
#     bot.alert(title='Warning', text='Script error found!')
#     df.at[line, 'Status CN21'] = 'Error'
#     df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
#     raise ValueError('\n\n------------- Error: -------------\n|> 2º Attribuition screen not found <|\n')

# press_key('stab', 1)
# bot.typewrite('ps0')
# press_key('enter', 1)

# # verificação sap
# if wait_event('images/ATTRIBUITION_2.png'):
#     pass
# else:
#     bot.alert(title='Warning', text='Script error found!')
#     df.at[line, 'Status CN21'] = 'Error'
#     df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')

# bot.typewrite('CJ02')
# press_key('enter', 1)
# bot.sleep(3)

# ===== MAIN =====

if __name__ == '__main__':
    wbs_element_creation()

    line = (df['Status CN21'] == 'Cadastrado').sum()
    repeat_qty = lp_qty - line

    bot.PAUSE = 0.35

    enter_cn21()
    diagram_creation()

    bot.alert(title='BotText', text='Program successfully completed')