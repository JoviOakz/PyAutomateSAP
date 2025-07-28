# ===== LIBRARIES =====

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re
from PIL import ImageEnhance, ImageFilter

# ===== CONSTANTS =====

KW = 25
ORIENTATION = 'pé'  # [deitado] | [pé]
ROTATION_ANGLE = 0  # [deitado -> 270] | [pé -> 0]

PDF_PATH = f'03 - PDF-Reader/LPs - KW{KW} - {ORIENTATION}.pdf'
OUTPUT_FILE = f'Open-LPs - {ORIENTATION}.xlsx'

DICTIONARY = {
    '—': '-',
    '~': '-',
    '‘': '-',
    '--': '-',
    '\'': '-',
    ',': '',
    'FL': 'P-',
    'F’': 'P',
    '_': '-',
    'o': '0',
    'O': '0',
    'LP"': 'LP-',
    'LP—': 'LP-',
    'LP_': 'LP-',
    'LP*': 'LP-',
    'LP<': 'LP-',
    'LP~': 'LP-',
    'LP0': 'LP-0',
    'LPO': 'LP-0',
    'LP-O': 'LP-0',
    'LP—O': 'LP-0',
    'LP_O': 'LP-0',
    'LP<O': 'LP-0',
    'LP~O': 'LP-0',
}

# ===== FUNCTIONS =====

def preprocess_image(image):
    image = image.convert('L')
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    return image

def preprocess_text(text, dictionary):
    for key, value in dictionary.items():
        text = text.replace(key, value)
    return text

def clear_string(s):
    return re.sub(r'[^LP0-9-]', '', s)

def find_flexible_lp(text):
    match = re.search(r'LP-\d{6}\b', text)
    if match:
        return match.group(0)

    matches = re.findall(r'L\s*P[\s\-_\W]*\d[\d\W]{4,10}', text, re.IGNORECASE)
    for match in matches:
        cleaned = re.sub(r'[^0-9]', '', match)
        if len(cleaned) == 6:
            lp = f'LP-{cleaned}'
            return lp
    return None

def extract_lps_from_pdf(pdf_path, rotation_angle=0, dictionary=None):
    images = convert_from_path(pdf_path)
    extracted_lps = []

    for i, image in enumerate(images, start=1):
        try:
            image = image.rotate(rotation_angle, expand=True)
            image = preprocess_image(image)

            text = pytesseract.image_to_string(image)
            text = preprocess_text(text, dictionary)

            lp = find_flexible_lp(text)

            if lp:
                extracted_lps.append(lp)
                print(f'[{i}/{len(images)}] LP encontrada: {lp}')
            else:
                print(f'[{i}/{len(images)}] LP não encontrada ❌')
                extracted_lps.append('')

        except Exception as e:
            print(f'[{i}/{len(images)}] Erro ao processar página ❌ - {e}')
            extracted_lps.append('')

    return extracted_lps

def save_lps_to_excel(lps, output_file):
    df = pd.DataFrame({'LP': lps})
    df.to_excel(output_file, index=False)
    print(f'\n✅ Arquivo salvo: {output_file}')

# ===== MAIN =====

def main():
    print(f'🔍 Extraindo LPs do arquivo: {PDF_PATH}')

    lps = extract_lps_from_pdf(
        PDF_PATH,
        rotation_angle=ROTATION_ANGLE,
        dictionary=DICTIONARY
    )

    save_lps_to_excel(lps, OUTPUT_FILE)

if __name__ == '__main__':
    main()