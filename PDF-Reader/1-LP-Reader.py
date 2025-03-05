import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re

def limpar_string(s):     
      return re.sub(r'[^LP0-9-]', '', s)

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
 
    extracted_text = []
    for i, image in enumerate(images):
        try:
            # CHANGE ROTATION
            rotated_image = image.rotate(0, expand=True)
 
            text = pytesseract.image_to_string(rotated_image)
            text = text.replace("—", "-").replace("~", "-")
            
            index0 = text.index("LP-")
            index1 = text.index(" ", index0+1)
    
            lp = text[index0:index1]
            extracted_text.append(lp.strip())
        except Exception as e:
            print("Error on page ", i+1, e)

    return extracted_text

# CHANGE PDF PATH
pdf_path = "PDF-Reader/LPs - KW05-2.pdf"
text = extract_text_from_pdf(pdf_path)
text = [limpar_string(s) for s in text]
print(text)

df = pd.DataFrame({"LP": text, "Status": ""})

df.to_excel("Open-LPs.xlsx")