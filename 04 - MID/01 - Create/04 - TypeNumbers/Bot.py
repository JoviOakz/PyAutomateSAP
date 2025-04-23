import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 1.25

bot.click(1802, 14)

excel_path = "Corpos BICO.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def exists_verification():
    bot.click(540, 450)
    bot.sleep(0.3)
    press_key('ctrla', 1)
    bot.typewrite(str(part_number))
    bot.sleep(0.3)
    press_key('enter', 1)

    bot.sleep(1)

    try:
        typeNumber_exist = list(bot.locateAllOnScreen('images/EXIST.png', grayscale=True, confidence=0.9))
        
        if typeNumber_exist:
            return True

    except Exception:
        print('Type number doesn\'t exist!')

line = 333
part_number_qty = 772

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    exist = False

    part_number = df.at[line, 'X'] # W -> 764 | X -> 772 | S -> 775 | K -> 29

    exist = exists_verification()

    if not exist:
        bot.click(1870, 1060)
        bot.sleep(0.3)
        bot.click(790, 450)
        bot.sleep(0.3)
        bot.typewrite(str(part_number))
        bot.sleep(0.3)
        press_key('enter', 1)
        bot.sleep(1.25)
        press_key('enter', 1)

    line += 1