# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd
import pyperclip as pc
import time
from pdf2image import convert_from_path
from PIL import ImageEnhance, ImageFilter
import pytesseract
import pdfplumber
import re

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.85

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = '../../01 - Excels/PO-Bot.xlsx'
df = pd.read_excel(
    EXCEL_PATH,
    engine='openpyxl',
    dtype={
        'Status': str
    }
)

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'sf6':
            bot.hotkey('shift', 'f6')
        elif key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def wait_event(img, region=None, timeout=7.5):
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            local = bot.locateOnScreen(img, region=region, grayscale=True, confidence=0.9)

            if local:
                return local
        except:
            pass

        time.sleep(0.5)
    return None

def preprocess_image(image):
    image = image.convert('L')
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    return image

def supplier_finder():
    image = convert_from_path(pdf_path)

    for i, image in enumerate(image):
        try:
            image = preprocess_image(image)
            pdf_text = pytesseract.image_to_string(image)

            if 'Costa Paula LTDA' in pdf_text:
                supplier = 'Retpress'
                return supplier
            else:
                text = re.search(r'\bGDS\b', pdf_text)

                if text:
                    supplier = text.group(0)
                    return supplier
                else:
                    return 'Erro'

        except Exception:
            print('Erro ao processar PDF')

def values_finder():
    with pdfplumber.open(pdf_path) as pdf:
        table = pdf.pages[0].extract_table()

        for _ in table:
            print(_)
        print('\n')

        qty = table[1][0].split(',')[0]
        description = table[1][2].split()[0]
        hr_piece = table[1][2].split('HR P/PÇ:')[1].split()[0]
        unit_value = table[1][4]
        item = table[1][6] 

    return qty, description, hr_piece, unit_value, item

def open_transaction():
    local = wait_event('images/SEARCH.png')
    
    if local:
        bot.click(local)
    else:
        bot.alert(title='Warning', text='Script error found!')
        df.at[line, 'Status'] = 'Error'
        df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
        raise ValueError('\n\n------------- Error: -------------\n|> Search not found <|\n')
    
    if 'LP' in debit_obj:
        bot.typewrite('CN22')
    else:
        bot.typewrite('IW32')

    press_key('enter', 1)

def open_debit_object():
    if 'LP' in debit_obj:
        # CAMINHO CN22
        pass
    else:
        if wait_event('images/IW32_1.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 1º IW32 screen not found <|\n')

        bot.typewrite(debit_obj)
        press_key('sf6', 1)

        if wait_event('images/IW32_2.png'):
            pass
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> 2º IW32 screen not found <|\n')

        press_key('ctrltab', 2)
        press_key('stab', 1)
        press_key('enter', 1)

        local = wait_event('images/MYIBUY.png', timeout=15)

        if local:
            bot.click(local)
        else:
            bot.alert(title='Warning', text='Script error found!')
            df.at[line, 'Status'] = 'Error'
            df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
            raise ValueError('\n\n------------- Error: -------------\n|> MyIbuy screen not found <|\n')

        bot.typewrite(f'{item[2]}h {supplier} sem mp')
        press_key('enter', 1)

# ===== PROGRAM CONFIGURATION =====

lp_qty = len(df['DebitObj'])
line = (df['Status'].notna()).sum()
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    pdf_path = ''
    supplier = ''
    item = []

    if (df['Status'].isna()).any():
        for _ in range(repeat_qty):
            debit_obj = str(df.at[line, 'DebitObj'])
            pdf_path = f'assets/{str(df.at[line, 'Orç'])}.pdf'
            supplier = supplier_finder()
            item = values_finder()

            print('Fornecedor: ' + supplier)
            print('Quantidade: ' + item[0])
            print('Norma: ' + item[1])
            print('HR P/PÇ: ' + item[2])
            print('Valor unitário: ' + item[3])
            print('Item: ' + item[4])

            if wait_event('images/MENU.png'):
                pass
            else:
                bot.alert(title='Warning', text='Script error found!')
                df.at[line, 'Status'] = 'Error'
                df.to_excel(EXCEL_PATH, index=False, engine='openpyxl')
                raise ValueError('\n\n------------- Error: -------------\n|> Menu screen not found <|\n')

            open_transaction()
            open_debit_object()
            
            line += 1

    bot.alert(title='BotText', text='Program successfully completed')