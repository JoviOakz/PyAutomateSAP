import pyautogui as bot
import pandas as pd
import pytesseract
import cv2
import numpy as np
from PIL import Image
import difflib

bot.FAILSAFE = True
bot.PAUSE = 0.35

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
    part_number = str(part_number).upper().replace('O', '0').replace(' ', '')

    screenshot = bot.screenshot(region=(600, 852, 542, 236))
    img_np = np.array(screenshot)

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)
    _, thresh = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    upscale = cv2.resize(thresh, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    final_image = Image.fromarray(upscale)

    custom_config = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    part_numbers_text = pytesseract.image_to_string(final_image, config=custom_config)

    part_numbers_list = [
        line.strip().upper().replace('O', '0').replace(' ', '')
        for line in part_numbers_text.splitlines() 
        if line.strip()
    ]

    print(part_numbers_list)

    if part_number in part_numbers_list:
        return part_numbers_list.index(part_number) + 1

    for idx, item in enumerate(part_numbers_list):
        ratio = difflib.SequenceMatcher(None, item, part_number).ratio()
        
        if ratio >= 0.96:
            print(f"Parecido com {part_number}: {item} (similaridade: {ratio:.2f})")
            return idx + 1

    return 0

part_number_qty = 2061
line = 0

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

        bot.sleep(0.5)

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
            
            bot.sleep(0.5)

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