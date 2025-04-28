import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.75

bot.click(1802, 14)

excel_path = "Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def register_verification():
    try:
        typeNumber_notExist = list(bot.locateAllOnScreen('images/NOTFIND.png', grayscale=True, confidence=0.9))
        
        if typeNumber_notExist:
            bot.click(1118, 832)
            bot.sleep(0.3)
            bot.click(1136, 1176)
            bot.sleep(0.3)
            bot.click(1136, 1084)
            bot.sleep(0.3)
            press_key('enter', 1)
            bot.sleep(0.3)
            bot.typewrite(str(part_number))
            bot.sleep(0.3)
            bot.click(1804, 14)
            
            return True
        
    except Exception:
        return False
        
part_number_qty = 772
line = 1

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    need_register = False

    part_number = df.at[line, 'X'] # W -> 764 | X -> 772 | S -> 775 | K -> 29

    bot.click(812, 830)
    bot.sleep(0.3)
    bot.typewrite(str(part_number))
    bot.sleep(1.5)

    need_register = register_verification()

    if not need_register:
        bot.click(812, 870)
        bot.sleep(0.3)
        press_key('tab', 1)
        bot.typewrite('01')
        bot.click(1620, 824)

    line += 1

bot.click(1852, 1074)