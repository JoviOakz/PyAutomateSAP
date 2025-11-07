# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.3

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def save():
    bot.click(1850, 946)
    bot.sleep(3)
    bot.click(1136, 664)

# ===== PROGRAM CONFIGURATION =====

set_qty = 1
line = 0
repeat_count = set_qty - line

# ===== MAIN =====

def main():
    global line

    for _ in range(repeat_count):
        norm = df.at[line, 'Norma']
        qty = df.at[line, 'Quantidade']
        name = df.at[line, 'Nome']
        life = df.at[line, 'Vida']
        interval = df.at[line, 'Intervalo']
        category = df.at[line, 'Categoria']
        sector = df.at[line, 'Setor']
        cc = df.at[line, 'Centro de custo']
        serie = 2

        for __ in range(qty):
            bot.click(1848, 942)
            bot.sleep(1.5)

            bot.click(532, 400)
            bot.typewrite(name)

            if serie < 10:
                bot.typewrite('0' + str(serie))
            else:
                bot.typewrite(str(serie))

            press_key('tab', 2)
            bot.typewrite(str(norm))
            bot.click(536, 536)
            press_key('tab', 1)

            if serie < 10:
                bot.typewrite('0' + str(serie))
            else:
                bot.typewrite(str(serie))

            press_key('tab', 1)
            bot.typewrite(str(life))
            press_key('tab', 1)
            bot.typewrite(str(interval))
            press_key('tab', 1)
            bot.typewrite(str(category))
            bot.click(560, 738)
            press_key('tab', 1)
            bot.typewrite(str(sector))
            bot.click(918, 738)
            press_key('tab', 1)
            bot.typewrite(str(cc))
            bot.click(1280, 738)
            press_key('tab', 2)

            if serie < 10:
                bot.typewrite('0' + str(serie))
            else:
                bot.typewrite(str(serie))

            bot.click(524, 958)
            press_key('tab', 1)
            press_key('space', 1)
            press_key('tab', 2)

            # =================================================

            bot.moveTo(580, 760, 0.15)
            bot.sleep(0.3)

            bot.scroll(-1676)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-1500)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')
            bot.hotkey('shift', 'tab')

            bot.scroll(-56)
            bot.click()

            bot.press('tab')
            bot.press('space')

            # =================================================

            bot.click(1660, 600)
            bot.scroll(-6000)
            bot.sleep(0.3)
            bot.click(1412, 766)
            bot.sleep(1)

            save()
            serie += 1

        line += 1

if __name__ == '__main__':
    main()