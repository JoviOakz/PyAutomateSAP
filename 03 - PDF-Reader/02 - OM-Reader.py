# ===== LIBRARIES =====

import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re

# ===== CONSTANTS =====

KW = 19
PDF_PATH = f"03 - PDF-Reader/OMs - KW{KW}.pdf"
OUTPUT_FILE = "Open-OMs.xlsx"
ROTATION_ANGLE = 0

DICTIONARY = {
    "—": "-",
    "‘": "",
}

# ===== FUNCTIONS =====

# --- REMOVE ALL EXTRA CHARACTERS ---
def limpar_string(s):    
    return re.sub(r'[^0-9]', '', s)

# --- EXTRACT ALL OM's FROM PDF ---
def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)

    extracted_text = []
    for i, image in enumerate(images):
        try:
            rotated_image = image.rotate(ROTATION_ANGLE, expand=True)
            text = pytesseract.image_to_string(rotated_image)

            for key, value in DICTIONARY.items():
                text = text.replace(key, value)

            text = text[20:]

            index0 = text.index("-")
            index1 = text.index(" ", index0 + 1)
            index2 = text.index(" ", index1 + 1)
            index3 = text.index(" ", index2 + 1)

            if len(text[index1:index2]) == 13 or len(text[index1:index2]) == 8:
                om = text[index1:index2]
            elif len(text[index2:index3]) == 13 or len(text[index2:index3]) == 8:
                om = text[index2:index3]
            else:
                om = '935'

            print(f"[{i + 1}/{len(images)}] OM encontrada: {om.strip()}")

            extracted_text.append(om.strip())
        except:
            print(f"[{i + 1}/{len(images)}] Erro ao processar página ❌")

    return extracted_text

# ===== MAIN =====

def main():
    print(f"🔍 Extraindo OMs do arquivo: {PDF_PATH}")
    
    text = extract_text_from_pdf(PDF_PATH)
    text = [limpar_string(s) for s in text]

    df = pd.DataFrame({"OM": text, "Status": ""})
    df.to_excel(OUTPUT_FILE, index=False)

    print(f"\n✅ Arquivo salvo: {OUTPUT_FILE}")

if __name__ == '__main__':
    main()