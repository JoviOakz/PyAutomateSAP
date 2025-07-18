# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pytesseract
import cv2
import numpy as np
from PIL import Image
import difflib

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.35

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data_reTest.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== FUNCTIONS =====

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

def normalize_ocr(text):
    text = text.upper().replace(' ', '')
    substitutions = {
        'O': '0',
        'I': '1',
        'L': '1',
        'Z': '2',
        'T': '7',
        'B': '8',
        'S': '5',
        '(': 'C',
        '"/': '7'
    }
    for wrong, right in substitutions.items():
        text = text.replace(wrong, right)
    return text

def find_position(part_number):
    part_number = str(part_number).upper().replace(' ', '')
    part_number_norm = normalize_ocr(part_number)

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
        line.strip().upper().replace(' ', '')
        for line in part_numbers_text.splitlines()
        if line.strip()
    ]

    for idx, item in enumerate(part_numbers_list):
        item_norm = normalize_ocr(item)
        ratio = difflib.SequenceMatcher(None, item_norm, part_number_norm).ratio()

        if ratio >= 0.9:
            return idx + 1

    return 0

# ===== PROGRAM CONFIGURATION =====

part_number_qty = 230
line = 0
repeat_count = part_number_qty - line

# ===== MAIN =====

def main():
    global line
    global part_number

    for _ in range(repeat_count):
        need_register = False

        part_number = df.at[line, 'Bico (normal final)']  # W -> 764 | X -> 772 | S -> 775 | K -> 29 | Bico (normal final) -> 2061

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

if __name__ == '__main__':
    main()