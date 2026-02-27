# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
from datetime import datetime, date

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 1.5

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../98 - Excels/Cadastro-LPs.xlsx'
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlf9':
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
        bot.sleep(1)

        bot.PAUSE = 0.4
        
        press_key('tab', 3)
        bot.sleep(1.25)
        press_key('ctrlf9', 1)
        bot.sleep(1.5)
        press_key('down', 2)
        press_key('tab', 1)
        press_key('down', 2)
        
        bot.PAUSE = 1.5

        press_key('ctrla', 1)
        bot.typewrite('68540028')

        bot.PAUSE = 0.4

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 1)
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 0.85

        press_key('tab', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        press_key('down', 1)

        prazo_final = datetime.strptime(df.at[line, 'Prazo Final'], '%d.%m.%Y').date()

        if date.today() > prazo_final:
            bot.typewrite(date.today().strftime('%d.%m.%Y'))
        else:
            bot.typewrite(df.at[line, 'Prazo Final'])

        press_key('down', 1)
        press_key('ctrla', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        press_key('down', 1)
        press_key('ctrla', 1)

        if date.today() > prazo_final:
            bot.typewrite(date.today().strftime('%d.%m.%Y'))
        else:
            bot.typewrite(df.at[line, 'Prazo Final'])
        
        bot.PAUSE = 0.4

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 3)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('tab', 1)
        press_key('down', 3)
        bot.sleep(0.65)

        bot.PAUSE = 0.85

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

        bot.PAUSE = 0.4

        press_key('stab', 4)
        press_key('enter', 1)
        bot.sleep(1.25)

        bot.PAUSE = 1.25

        press_key('tab', 1)





        

        # if valor for centro de custo no caso 6 digitos -> ZPS001
        # if valor começa com BM -> ZPS007 e esquema 07
        
        # --
        bot.typewrite('ZPS001')
        # --
        
        # calculo objeto de liquidação (ex: ZPS007)
        
        press_key('tab', 1)
        bot.typewrite(str(df.at[line, 'Esquema de Alocação']))









        press_key('f3', 1)
        bot.sleep(0.85)
        press_key('tab', 1)
        bot.typewrite(str(df.at[line, 'Objeto de Liquidação']))
        press_key('f3', 1)
        bot.sleep(0.85)
        press_key('f3', 1)
        bot.sleep(0.85)
        press_key('sf1', 1)
        press_key('down', 1)
        press_key('tab', 2)
        bot.typewrite('68540028')
        press_key('down', 1)
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        press_key('down', 1)

        # --
        bot.typewrite(date.today().strftime('%d.%m.%Y'))
        # --

        # if data_atual > data_entrega:
        #     bot.typewrite(date.today().strftime('%d.%m.%Y'))
        # else:
        #     bot.typewrite(data_entrega)

        # fazer processo mouse para liberar

        # press_key('ctrls', 1)
        bot.sleep(2)
        
        line += 1

    bot.sleep(2)

def diagram_creation():
    global line

    bot.PAUSE = 0.4

    press_key('stab', 1)
    press_key('left', 7)
    bot.typewrite('/ncn21')
    press_key('enter', 1)
    bot.sleep(3)

    bot.PAUSE = 0.4

    press_key('right', 3)
    press_key('tab', 1)
    bot.typewrite('BP01')
    press_key('tab', 1)
    bot.sleep(1.25)
    bot.typewrite('6854')
    press_key('tab', 1)
    bot.sleep(1.25)
    bot.typewrite('I33')
    bot.sleep(1.25)

    bot.PAUSE = 1.25

    for _ in range(repeat_qty):
        press_key('enter', 1)
        bot.typewrite(df.at[line, 'Elemento PEP'].replace('-', ''))
        press_key('enter', 1)
        bot.sleep(1)

        bot.PAUSE = 0.4

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

        bot.PAUSE = 0.4

        press_key('tab', 2)
        press_key('right', 3)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('tab', 4)

        bot.PAUSE = 0.85
        
        text = 'HEIJUNKA\n' + str(df.at[line, 'Part Number']) + ' - ' + df.at[line, 'Denominação'] + '\nResp. Yesica Gonzalez'
        bot.typewrite(text)
        press_key('ctrlsf12', 1)
        bot.sleep(1.25)
        press_key('ctrls', 1)
        bot.sleep(2)

        line += 1

# ===== PROGRAM CONFIGURATION =====

lp_qty = 2
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    wbs_element_creation()
    diagram_creation()

    bot.alert(title='BotText', text='Programa encerrado!')