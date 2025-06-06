import pyautogui as bot
import pandas as pd
import pytesseract

bot.FAILSAFE = True
bot.PAUSE = 0.5

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
            bot.click(1126, 834)
            bot.sleep(0.3)
            bot.click(1136, 1176)
            bot.sleep(0.3)
            press_key('enter', 1)
            bot.sleep(0.3)
            bot.typewrite(str(part_number))

            bot.sleep(0.3)
            bot.click(1136, 1176)
            
            return True
        
    except Exception:
        return False
    
def find_position(part_number):
    part_numbers_image = bot.screenshot(region=(600, 852, 542, 236))
    part_numbers_text = pytesseract.image_to_string(part_numbers_image)
    part_numbers_list = [line.strip() for line in part_numbers_text.splitlines() if line.strip()]
    
    if part_number in part_numbers_list:
        return part_numbers_list.index(part_number) + 1
    return 0

part_number_qty = 2061
line = 7

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    need_register = False

    part_number = df.at[line, 'Bico (normal final)'] # Bico (normal final) -> 2061

    if str(part_number).startswith('433'):
        part_number = ('0' + str(part_number))

    bot.click(812, 830)
    bot.sleep(0.3)

    bot.typewrite(str(part_number))

    bot.sleep(1.25)

    need_register = register_verification()

    if not need_register:
        click_coords = {
            1: (874, 870),
            2: (874, 922),
            3: (874, 970),
            4: (874, 1020),
            5: (874, 1070)
        }

        bot.moveTo(964, 920)
        bot.scroll(1000)
        bot.moveTo(400, 720)

        posicao = find_position(part_number)

        if posicao in click_coords:
            x, y = click_coords[posicao]
            bot.click(x, y)
            press_key('tab', 1)
            bot.typewrite('01')
            press_key('tab', 2)
            press_key('space', 1)
        else:
            bot.moveTo(964, 920)
            bot.scroll(-290)
            bot.moveTo(400, 720)

            posicao = find_position(part_number)

            if posicao in click_coords:
                x, y = click_coords[posicao]
                bot.click(x, y)
                press_key('tab', 1)
                bot.typewrite('01')
                press_key('tab', 2)
                press_key('space', 1)
            else:
                bot.moveTo(964, 920)
                bot.scroll(-290)
                bot.moveTo(400, 720)

                posicao = find_position(part_number)

                if posicao in click_coords:
                    x, y = click_coords[posicao]
                    bot.click(x, y)
                    press_key('tab', 1)
                    bot.typewrite('01')
                    press_key('tab', 2)
                    press_key('space', 1)
                else:
                    bot.click(1112, 826)

                    bot.click(1136, 1176)
                    bot.sleep(0.3)
                    press_key('enter', 1)
                    bot.sleep(0.3)
                    bot.typewrite(str(part_number))

                    bot.sleep(0.3)
                    bot.click(1136, 1176)

    line += 1

bot.sleep(0.3)
bot.click(1852, 1074)