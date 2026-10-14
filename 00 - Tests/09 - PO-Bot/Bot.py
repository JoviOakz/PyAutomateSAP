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

search_color = (255, 150, 50)

# ===== DICTIONARY =====

CART_1 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\CART_1.png'
CART_2 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\CART_2.png'
CART_3 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\CART_3.png'
IW32_1 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\IW32_1.png'
IW32_2 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\IW32_2.png'
MENU = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\MENU.png'
MYIBUY = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\MYIBUY.png'
SAP_1 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\SAP_1.png'
SAP_2 = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\SAP_2.png'
SEARCH = r'C:\Users\Mao8ct\Desktop\PyAutomateSAP\00 - Tests\09 - PO-Bot\images\SEARCH.png'

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlf':
            bot.hotkey('ctrl', 'f')
        elif key == 'sf6':
            bot.hotkey('shift', 'f6')
        elif key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def wait_event(img, region=None, timeout=5):
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

def sap_start():
    press_key('winr', 1)
    bot.typewrite('saplogon')
    press_key('enter', 1)

    if wait_event(SAP_1):
        pass
    else:
        close_sap()
        raise Exception('Error')

    bot.sleep(1)
    press_key('stab', 1)
    bot.typewrite('ps0')
    press_key('enter', 1)

    if wait_event(SAP_2):
        pass
    else:
        close_sap()
        raise Exception('Error')

    bot.sleep(1)

    bot.typewrite('IW32')

    press_key('enter', 1)

def close_sap():
    press_key('winr', 1)
    bot.typewrite('cmd /c taskkill /f /im saplogon.exe')
    press_key('enter', 1)

def search_click():
    for _ in range(20):
        screenshot = bot.screenshot()
        width, height = screenshot.size

        for y in range(0, height, 1):
            for x in range(0, width, 1):
                if screenshot.getpixel((x, y)) == search_color:
                    bot.click(x, y)
                    return True
    return False
 
def open_debit_object():
    if 'LP' in debit_obj:
        # CAMINHO CN22
        pass
    else:
        if wait_event(IW32_1):
            pass
        else:
            close_sap()
            raise Exception('Error')

        bot.typewrite(debit_obj)
        press_key('sf6', 1)

        if wait_event(IW32_2):
            pass
        else:
            close_sap()
            raise Exception('Error')

        press_key('ctrltab', 2)
        press_key('stab', 1)
        press_key('enter', 1)

def search_usi():
        if wait_event(MYIBUY, timeout=15):
            pass
        else:
            close_sap()
            raise Exception('Error')

        press_key('ctrltab', 1)
        press_key('tab', 2)
        pc.copy(f'Usin.{item[2]}h({supplier}) - (Geral) - Sem mp')
        press_key('ctrlv', 1)
        press_key('enter', 1)
        bot.sleep(1)
        press_key('ctrlf', 1)
        press_key('ctrlv', 1)
        press_key('tab', 2)
        press_key('enter', 1)
        search_click()

        if wait_event(CART_1):
            pass
        else:
            close_sap()
            raise Exception('Error')
            
        cart_image = bot.locateOnScreen(CART_1)
        bot.click(bot.center(cart_image))
        bot.typewrite('3')
        bot.sleep(1)
        cart_image = bot.locateOnScreen(CART_2)
        bot.click(bot.center(cart_image))
        
        if wait_event(CART_3):
            pass
        else:
            close_sap()
            raise Exception('Error')
        
        cart_image = bot.locateOnScreen(CART_3)
        bot.click(bot.center(cart_image))

# ===== PROGRAM CONFIGURATION =====

EXCEL_PATH = '../../01 - Excels/PO-Bot.xlsx'
df = pd.read_excel(
    EXCEL_PATH,
    engine='openpyxl',
    dtype={
        'Status': str
    }
)

# lp_qty = len(df['DebitObj'])
# line = (df['Status'].notna()).sum()
# repeat_qty = lp_qty - line
lp_qty = 1
line = 0

# ===== MAIN =====

def main():
    supplier = ''
    item = []

    supplier = supplier_finder()
    item = values_finder()

    u_value = str(round(float(item[3]) * float(item[5]), 2))
    u_value = u_value.replace('.', ',')
    t_value = str(round(float(item[2]) * float(item[3]) * float(item[5]), 2))
    t_value = t_value.replace('.', ',')

    print('\n==========================================================\n')

    print(f'''Fornecedor: {supplier}
    Orç: XYZ - ITEM 0{item[0]}
    LP/OM: XYZ
    Descrição: {item[1]}
    Qtd.: {item[2]} peças
    Prazo: DD/MM/YYYY
    Responsável

    Valor unitário (cálculo): {f'R${u_value}'}
    Valor total (pytesseract): {'R$' + item[4]}
    Valor total (cálculo): {f'R${t_value}'}
    Tempo de usinagem: {item[5]} horas p/pç''')

    print('\n==========================================================\n')

    # sap_start()

    debit_obj = str(df.at[line, 'DebitObj'])
    pdf_path = f'assets/gds.pdf'
    # pdf_path = f'assets/{str(df.at[line, 'Orç'])}.pdf'

    print('Fornecedor: ' + supplier)
    print('Quantidade: ' + item[0])
    print('Norma: ' + item[1])
    print('HR P/PÇ: ' + item[2])
    print('Valor unitário: ' + item[3])
    print('Item: ' + item[4])

    # open_debit_object()
    search_usi()
    # confirm_po()
    
    bot.alert(title='BotText', text='Program successfully completed')

if __name__ == '__main__':
    main()