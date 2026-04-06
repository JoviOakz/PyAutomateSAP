# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
from datetime import date

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.5

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/LP-Register.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

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
        elif key == 'ctrlsf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        else:
            bot.press(key)

def wbs_element_creation():
    global line

    for _ in range(repeat_qty):
        bot.typewrite(df.at[line, 'Elemento PEP'])
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.75
        
        press_key('ctrla', 1)

        part_number = str(df.at[line, 'Part Number']).replace(' ', '')

        pc.copy(str(df.at[line, 'Denominação']))

        if not part_number.isalpha():
            if len(part_number) >= 10:
                full_text = str(df.at[line, 'Part Number']) + str(df.at[line, 'Denominação'])
                if len(full_text) <= 37:
                    bot.typewrite(str(df.at[line, 'Part Number']) + ' - ')
                    press_key('ctrlv', 1)
                else:
                    bot.typewrite(str(df.at[line, 'Part Number']))
            else:
                press_key('ctrlv', 1)
        else:
            press_key('ctrlv', 1)

        bot.sleep(1.25)
        press_key('tab', 3)
        bot.sleep(0.85)

        if not part_number.isalpha():
            if len(part_number) >= 10:
                full_text = str(df.at[line, 'Part Number']) + str(df.at[line, 'Denominação'])
                if len(full_text) <= 37:
                    bot.typewrite(str(df.at[line, 'Part Number']) + ' - ')
                    press_key('ctrlv', 1)
                else:
                    bot.typewrite(str(df.at[line, 'Part Number']))
            else:
                press_key('ctrlv', 1)
        else:
            press_key('ctrlv', 1)

        bot.sleep(1.5)
        press_key('ctrlf9', 1)
        bot.sleep(1.5)
        press_key('down', 2)
        press_key('tab', 1)
        press_key('down', 2)
        
        bot.PAUSE = 1.5

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

        bot.PAUSE = 0.75

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 4)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('tab', 1)
        press_key('down', 3)
        bot.sleep(0.65)

        bot.PAUSE = 1.15

        bot.typewrite(df.at[line, 'Entregar para'])
        bot.sleep(0.85)
        press_key('tab', 2)
        bot.typewrite(str(df.at[line, 'Quantidade']))
        press_key('tab', 1)
        bot.typewrite('PC')
        press_key('tab', 1)
        press_key('down', 1)
        bot.typewrite(str(df.at[line, 'Custo estimado']))
        press_key('tab', 1)
        bot.typewrite('BRL')
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.75

        press_key('stab', 4)
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 1.5

        press_key('tab', 1)

        liquidation_object = str(df.at[line, 'Objeto de Liquidação']).strip()
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
        bot.sleep(0.85)
        press_key('tab', 1)
        bot.typewrite(str(df.at[line, 'Objeto de Liquidação']))
        press_key('f3', 1)
        bot.sleep(0.85)
        press_key('f3', 1)
        bot.sleep(0.85)
        press_key('sf1', 1)
        bot.sleep(0.5)

        bot.PAUSE = 0.85

        press_key('tab', 1)
        press_key('down', 1)
        press_key('tab', 2)
        bot.typewrite(iss_dept)
        press_key('down', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        press_key('down', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))

        bot.PAUSE = 1

        bot.click(150, 14)
        bot.click(206, 106)
        bot.click(450, 106)
        bot.sleep(1.25)

        press_key('ctrls', 1)
        df.at[line, 'Status CJ02'] = 'Cadastrada'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        bot.sleep(2.5)
        
        line += 1

    bot.sleep(2)

def diagram_creation():
    global line

    for _ in range(repeat_qty):
        press_key('enter', 1)
        bot.typewrite(df.at[line, 'Elemento PEP'].replace('-', ''))
        press_key('enter', 1)
        bot.sleep(1)

        bot.PAUSE = 0.75

        press_key('tab', 2)
        press_key('right', 1)
        press_key('enter', 1)
        bot.sleep(0.85)
        press_key('tab', 2)
        bot.typewrite(df.at[line, 'Elemento PEP'].replace('-', ''))
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('tab', 1)
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.75

        press_key('tab', 2)
        press_key('right', 3)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('tab', 4)

        bot.PAUSE = 1.15
        
        press_key('ctrla', 1)
        text_partnumber = 'HEIJUNKA\n' + str(df.at[line, 'Part Number']) + ' - '
        text_resp = '\nResp. ' + str(df.at[line, 'Responsável'])
        bot.typewrite(text_partnumber)
        press_key('ctrlv', 1)
        bot.typewrite(text_resp)
        bot.sleep(1)
        press_key('ctrlsf12', 1)
        bot.sleep(2)
        press_key('ctrls', 1)
        df.at[line, 'Status CN21'] = 'Cadastrada'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        bot.sleep(2.5)

        line += 1

# ===== PROGRAM CONFIGURATION =====

lp_qty = 128
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    wbs_element_creation()

    bot.PAUSE = 0.75

    press_key('stab', 1)
    press_key('left', 7)
    bot.typewrite('/ncn21')
    press_key('enter', 1)
    bot.sleep(3)

    bot.PAUSE = 0.75

    press_key('right', 3)
    press_key('tab', 1)
    bot.typewrite('BP01')
    press_key('tab', 1)
    bot.sleep(1.25)
    bot.typewrite('6854')
    press_key('tab', 1)
    bot.sleep(1.25)

    line = 0
    repeat_qty = lp_qty - line

    if str(df.at[line, 'Responsável']) == 'Yesica Gonzalez':
        bot.typewrite('I33')
    elif str(df.at[line, 'Responsável']) == 'Rodrigo Melo':
        bot.typewrite('I49')
    else:
        bot.typewrite('I39')
    
    bot.sleep(1.25)

    bot.PAUSE = 1.25

    diagram_creation()

    bot.alert(title='BotText', text='Programa encerrado!')