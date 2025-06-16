# ===== LIBRARIES =====

import win32com.client as win32
import pyautogui as bot

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.01

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r"C:\Users\Mao8ct\Desktop\PyAutomateSAP\05 - WorkedHourPointing\02 - OM's\ApontamentoFrança.xlsx")
ws = wb.Sheets('Sheet1')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

# ===== MAIN =====

def main():
    try:
        bot.sleep(0.5)
        press_key('tab', 1)

        ws.Range('A2:B23').Copy()
        bot.hotkey('ctrl', 'v')
        bot.sleep(0.5)

        ws.Range('C2:E23').Copy()
        press_key('tab', 3)
        bot.hotkey('ctrl', 'v')
        bot.sleep(0.5)

        ws.Range('F2:J23').Copy()
        press_key('tab', 5)
        bot.hotkey('ctrl', 'v')
        bot.sleep(0.5)
        
        press_key('shtab', 2)

        for __ in range(21):
            press_key('space', 1)
            press_key('down', 1)
        press_key('space', 1)

        press_key('tab', 1)

        for __ in range(21):
            press_key('space', 1)
            press_key('up', 1)
        press_key('space', 1)

        press_key('tab', 10)

        for __ in range(21):
            press_key('space', 1)
            press_key('down', 1)
        press_key('space', 1)

    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

if __name__ == '__main__':
    main()



# ===== CALCULO A SER FEITO =====

# valor total mês / qtd total de ordens
# 20.000 ÷ 22 = R$909,09 por OT

# qtd em reais por ordem / qtd reais por hora
# 909,09 ÷ 120 = 7,58 horas por OT

# França precisa de 167 horas para bater sua meta de 20.000 no mês