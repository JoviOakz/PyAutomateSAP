# ===== LIBRARIES =====

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re

# ===== CONSTANTS =====

ROTATION_ANGLE = 270 # [deitado -> 270] | [pé -> 0]

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

# --- REMOVE ALL EXTRA CHARACTERS ---
def clear_string(s):     
    return re.sub(r'[^LP0-9-]', '', s)

# --- ADJUST LP's LENGTH ---
def normalize_lp(t):
    if len(t) > 9:
        return t[:9]
    elif len(t) == 8:
        return t[:3] + '0' + t[3:]
    return t

# --- FIX LP's PRE ISSUES ---
def preprocess_text(text, dictionary):
    for key, value in dictionary.items():
        text = text.replace(key, value)
    return text

# --- EXTRACT ALL LP's FROM PDF ---
def extract_lps_from_pdf(pdf_path, rotation_angle=0, dictionary=None):
    images = convert_from_path(pdf_path)
    extracted_lps = []

    for i, image in enumerate(images, start=1):
        try:
            image = image.rotate(rotation_angle, expand=True)
            
            text = pytesseract.image_to_string(image)

            preprocess_text(text, dictionary)

            index0 = text.index('LP')
            index1 = text.index(' ', index0+1)
    
            lp = text[index0:index1]
            extracted_lps.append(lp.strip())

            print(f'[{i}/{len(images)}] LP encontrada: {lp[:9]}')
            
        except Exception:
            print(f'[{i}/{len(images)}] LP não encontrada ❌')

    return extracted_lps

# --- SAVE EXTRACTED LP's IN EXCEL ---
def save_lps_to_excel(lps, output_file):
    cleaned_lps = [normalize_lp(clear_string(lp)) for lp in lps]
    df = pd.DataFrame({'LP': cleaned_lps, 'Status': ''})
    df.to_excel(output_file, index=False)
    print(f'\n✅ Arquivo salvo: {output_file}')

# ===== MAIN =====

def main():
    kw = 19
    orientation = 'deitado' # [deitado] | [pé]

    pdf_path = f'03 - PDF-Reader/LPs - KW{kw} - {orientation}.pdf'
    output_file = f'Open-LPs - {orientation}.xlsx'

    print(f'🔍 Extraindo LPs do arquivo: {pdf_path}')

    lps = extract_lps_from_pdf(
        pdf_path, 
        rotation_angle=ROTATION_ANGLE, 
        dictionary=DICTIONARY
    )

    save_lps_to_excel(lps, output_file)

if __name__ == '__main__':
    main()