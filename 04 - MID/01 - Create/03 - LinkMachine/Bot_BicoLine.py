import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 1

bot.click(1802, 14)

excel_path = "Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

part_number_coords = (596, 850, 454, 242)

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
            press_key('enter', 1)
            bot.sleep(0.3)
            
            if str(part_number).startswith('433'):
                bot.typewrite('0' + str(part_number))
            else:
                bot.typewrite(str(part_number))

            bot.sleep(0.3)
            bot.click(1136, 1176)
            
            return True
        
    except Exception:
        return False

part_number_qty = 2061
line = 2019

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    need_register = False

    part_number = df.at[line, 'Bico (normal final)'] # Bico (normal final) -> 2061

    bot.click(812, 830)
    bot.sleep(0.3)

    if str(part_number).startswith('433'):
        bot.typewrite('0' + str(part_number))
    else:
        bot.typewrite(str(part_number))
    
    bot.sleep(1.5)

    need_register = register_verification()

    if not need_register:
        part_number_image = bot.screenshot(region=(596, 822, 104, 24))

        bot.moveTo(852, 872, 0.15)
        bot.sleep(0.15)
        bot.scroll(1000)
        bot.sleep(0.15)
        bot.moveTo(852, 700, 0.15)

        try:
            part_number_found = list(bot.locateOnScreen(part_number_image, grayscale=True, confidence=0.9, region=part_number_coords))
            
            if part_number_found:
                x, y = bot.locateCenterOnScreen(part_number_image, grayscale=True, confidence=0.9, region=part_number_coords)

                bot.click(x, y)
                press_key('tab', 1)
                bot.typewrite('01')
                bot.click(1620, 824)

        except Exception:
            bot.moveTo(852, 872, 0.15)
            bot.sleep(0.15)
            bot.scroll(-400)
            bot.sleep(0.15)
            bot.moveTo(852, 700, 0.15)
            
            try:
                part_number_found = list(bot.locateOnScreen(part_number_image, grayscale=True, confidence=0.9, region=part_number_coords))
                
                if part_number_found:
                    x, y = bot.locateCenterOnScreen(part_number_image, grayscale=True, confidence=0.9, region=part_number_coords)

                    bot.click(x, y)
                    press_key('tab', 1)
                    bot.typewrite('01')
                    bot.click(1620, 824)

            except Exception:
                bot.click(1118, 832)
                bot.sleep(0.3)
                bot.click(1136, 1176)
                bot.sleep(0.3)
                press_key('enter', 1)
                bot.sleep(0.3)

                if str(part_number).startswith('433'):
                    bot.typewrite('0' + str(part_number))
                else:
                    bot.typewrite(str(part_number))

                bot.sleep(0.3)
                bot.click(1136, 1176)

    line += 1

bot.click(1852, 1074)