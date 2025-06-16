# ===== LIBRARIES =====

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re
from PIL import Image, ImageEnhance, ImageFilter

# ===== CONSTANTS =====

KW = 21
ROTATION_ANGLE = 0

PDF_PATH = f'03 - PDF-Reader/OMs - KW{KW}.pdf'
OUTPUT_FILE = 'Open-OMs.xlsx'

DICTIONARY = {
    '—': '-',
    '~': '-',
    '‘': '',
    '\'': '',
    '_': '-',
    'o': '0',
    'O': '0',
    ',': '',
}

# ===== FUNCTIONS =====

def limpar_string(s):    
    return re.sub(r'[^0-9]', '', s)

def preprocess_image(image):
    image = image.convert('L')  # escala de cinza
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)

    return image

def preprocess_text(text, dictionary):
    for key, value in dictionary.items():
        text = text.replace(key, value)

    return text

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
    extracted_oms = []

    for i, image in enumerate(images):
        try:
            rotated_image = image.rotate(ROTATION_ANGLE, expand=True)
            preprocessed_image = preprocess_image(rotated_image)

            text = pytesseract.image_to_string(preprocessed_image, config='--psm 6')
            text = preprocess_text(text, DICTIONARY)

            match = re.search(r'\b\d{5,}\b', text)

            if match:
                om = match.group(0)
                om_limpa = limpar_string(om)
                extracted_oms.append(om_limpa)
                print(f'[{i + 1}/{len(images)}] OM encontrada: {om_limpa}')
            else:
                print(f'[{i + 1}/{len(images)}] Nenhuma OM encontrada ❌')
                extracted_oms.append('')

        except Exception as e:
            print(f'[{i + 1}/{len(images)}] Erro ao processar página ❌ - {e}')
            extracted_oms.append('')

    return extracted_oms

# ===== MAIN =====

def main():
    print(f'🔍 Extraindo OMs do arquivo: {PDF_PATH}')
    
    oms = extract_text_from_pdf(PDF_PATH)

    df = pd.DataFrame({'OM': oms, 'Status': ''})
    df.to_excel(OUTPUT_FILE, index=False)

    print(f'\n✅ Arquivo salvo: {OUTPUT_FILE}')

if __name__ == '__main__':
    main()