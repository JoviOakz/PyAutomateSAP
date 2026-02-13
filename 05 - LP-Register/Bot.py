# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
from datetime import date

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
        else:
            bot.press(key)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 1
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_qty):
        bot.typewrite(df.at[line, 'Elemento PEP'])
        press_key('enter', 1)
        bot.sleep(1)

        part_number = df.at[line, 'Part Number']
        description = df.at[line, 'Denominação']
        lp_title = f'{part_number} - {description}'
        bot.typewrite(lp_title)

        bot.PAUSE = 0.25
        
        press_key('tab', 3)
        bot.sleep(1.15)
        bot.typewrite(lp_title)
        bot.sleep(1.15)
        press_key('ctrlf9', 1)
        bot.sleep(1.5)
        press_key('down', 2)
        press_key('tab', 1)
        press_key('down', 2)
        
        bot.PAUSE = 1.5

        press_key('ctrla', 1)
        bot.typewrite('68540028')

        bot.PAUSE = 0.25

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 1)
        press_key('enter', 1)
        bot.sleep(1.15)

        bot.PAUSE = 0.85

        press_key('tab', 1)
        bot.typewrite(date.today().strftime("%d.%m.%Y"))
        press_key('down', 1)

        # --
        bot.typewrite(date.today().strftime("%d.%m.%Y"))
        # --

        # if data_atual > data_entrega:
        #     bot.typewrite(date.today().strftime("%d.%m.%Y"))
        # else:
        #     bot.typewrite(data_entrega)

        press_key('down', 1)
        press_key('ctrla', 1)
        bot.typewrite(date.today().strftime("%d.%m.%Y"))
        press_key('down', 1)
        press_key('ctrla', 1)

        # --
        bot.typewrite(date.today().strftime("%d.%m.%Y"))
        # --

        # if data_atual > data_entrega:
        #     bot.typewrite(date.today().strftime("%d.%m.%Y"))
        # else:
        #     bot.typewrite(data_entrega)
        
        bot.PAUSE = 0.25

        press_key('up', 3)
        press_key('stab', 1)
        press_key('right', 3)
        press_key('enter', 1)
        bot.sleep(1.15)
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

        bot.PAUSE = 0.25

        press_key('stab', 4)
        press_key('enter', 1)
        bot.sleep(1.15)

        bot.PAUSE = 0.85

        press_key('tab', 1)
        
        # --
        bot.typewrite('ZPS007')
        # --
        
        # calculo objeto de liquidação (ex: ZPS007)
        
        press_key('tab', 1)
        bot.typewrite(str(df.at[line, 'Esquema de Alocação']))
        press_key('f3', 1)
        bot.sleep(0.65)
        press_key('f3', 1)
        bot.sleep(0.65)
        press_key('sf1', 1)
        press_key('down', 1)
        press_key('tab', 1)
        press_key('down', 1)
        press_key('tab', 1)
        bot.typewrite('68540028')
        
        line += 1

if __name__ == '__main__':
    main()
    bot.alert(title='BotText', text='Programa encerrado!')