from pdf2image import convert_from_path
from PIL import ImageEnhance, ImageFilter
import pytesseract
import pdfplumber
import re

PDF_PATH = r'GDS - Orç.06572-26.pdf'

def preprocess_image(image):
    image = image.convert('L')
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    return image

def supplier_finder():
    image = convert_from_path(PDF_PATH)

    for i, image in enumerate(image):
        try:
            image = preprocess_image(image)
            text = pytesseract.image_to_string(image)
            supplier = re.search(r'\bGDS\b', text)

            if supplier:
                text = supplier.group(0)
                return text
            else:
                print('Erro ao encontrar o fornecedor')

        except Exception:
            print('Erro ao processar PDF')

def values_finder():
    with pdfplumber.open(PDF_PATH) as pdf:
        table = pdf.pages[0].extract_table()

        number = table[1][0]
        text = table[1][1]
        qty = table[1][2]
        usi_time = re.sub(r'[^0-9,]', '', table[1][3])
        om = re.sub(r'[^0-9]', '', table[5][1])

    return number, text, qty, usi_time, om

def main():
    supplier = ''
    item = []

    supplier = supplier_finder()
    item = values_finder()

    print('\n==========================================================\n')

    print(f'''Fornecedor: {supplier}
    Orç: 06572-26 - ITEM 0{item[0]}
    LP/OM: {item[4]}
    Descrição: {item[1]}
    Qtd.: {item[2]} peças
    Prazo: 07/08/2026
    Thais Fischer

    Tempo de usinagem: {item[3]} horas p/pç''')

    print('\n==========================================================\n')

if __name__ == '__main__':
    main()