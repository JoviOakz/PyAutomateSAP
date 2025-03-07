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
            rotated_image = image.rotate(90, expand=True)
 
            text = pytesseract.image_to_string(rotated_image)

            dictionary = {
                "—":"-",
                "~":"-",
                ",":"",
                "FL":"P-",
                "F’":"P",
                "_":"-",
                "o":"0",
                "O":"0",
                "LP0":"LP-0",
            }
            
            for key, value in dictionary.items():
                text = text.replace(key, value)
            
            index0 = text.index("LP")
            index1 = text.index(" ", index0+1)
    
            lp = text[index0:index1]
            extracted_text.append(lp.strip())
        except Exception as e:
            print("Error on page", i+1, e)

    return extracted_text

# CHANGE PDF PATH
pdf_path = "PDF-Reader/LPs - KW05.pdf"
text = extract_text_from_pdf(pdf_path)
text = [limpar_string(s) for s in text]

text = [
    t[:9] if len(t) > 9 else t[:3] + "0" + t[3:] if len(t) == 8 else t
    for t in text
]

print(text)

df = pd.DataFrame({"LP": text, "Status": ""})

df.to_excel("Open-LPs.xlsx")