import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re

def limpar_string(s):     
      return re.sub(r'[^0-9]', '', s)

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
 
    extracted_text = []
    for i, image in enumerate(images):
        try:
            rotated_image = image.rotate(0, expand=True)
 
            text = pytesseract.image_to_string(rotated_image)

            index0 = text.index("-")
            index1 = text.index(" ", index0+1)
            index2 = text.index(" ", index1+1)

            om = text[index1:index2]

            extracted_text.append(om.strip())
        except:
            print("Error on page ", i+1)

    return extracted_text
 
pdf_path = "PDF-Reader/OMs - KW05.pdf"
text = extract_text_from_pdf(pdf_path)
text = [limpar_string(s) for s in text]
print(text)

df = pd.DataFrame({"OM": text, "Status": ""})

df.to_excel("Open-OMs.xlsx")