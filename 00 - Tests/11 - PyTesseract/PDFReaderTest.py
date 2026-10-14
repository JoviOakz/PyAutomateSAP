from pdf2image import convert_from_path
from PIL import ImageEnhance, ImageFilter
import pytesseract
import pdfplumber
import re

PDF_PATH = r'eletrospark.pdf'

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
            supplier = re.search(r'\b(GDS|Eletrospark)\b', text, re.IGNORECASE)

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
        h_value = re.sub(r'[^\d,\.]', '', table[1][4])
        h_value = h_value.replace(',', '.')
        t_value = re.sub(r'[^\d,\.]', '', table[1][5])
        usi_time = re.sub(r'[^0-9,]', '', table[1][3])
        usi_time = usi_time.replace(',', '.')

    return number, text, qty, h_value, t_value, usi_time

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

if __name__ == '__main__':
    main()