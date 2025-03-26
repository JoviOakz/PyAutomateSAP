# LIBRARIES
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
import re
 
# REMOVE ALL EXTRA CHARACTERS
def limpar_string(s):    
      return re.sub(r'[^0-9]', '', s)
 
# CATCH ALL OM's IN PDF
def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
 
    extracted_text = []
    for i, image in enumerate(images):
        try:
            # CHANGE THE PDF ROTATION
            rotated_image = image.rotate(0, expand=True)
 
            text = pytesseract.image_to_string(rotated_image)
 
            dictionary = {
                "—":"-",
                "‘":"",
            }
            
            for key, value in dictionary.items():
                text = text.replace(key, value)

            text = text[20:]

            index0 = text.index("-")
            index1 = text.index(" ", index0+1)
            index2 = text.index(" ", index1+1)
            index3 = text.index(" ", index2+1)
 
            if len(text[index1:index2]) == 13 or len(text[index1:index2]) == 8:
                om = text[index1:index2]
            elif len(text[index2:index3]) == 13 or len(text[index2:index3]) == 8:
                om = text[index2:index3]
            else:
                om = '935'

            print(om)
 
            extracted_text.append(om.strip())
        except:
            print("Error on page ", i+1)
 
    return extracted_text
 
# PDF PATH
pdf_path = "PDF-Reader/OMs - KW07.pdf"

# FILTER OM's CORRECTLY
text = extract_text_from_pdf(pdf_path)
text = [limpar_string(s) for s in text]
print(text)
 
# EXTRACT TO EXCEL
df = pd.DataFrame({"OM": text, "Status": ""})
df.to_excel("Open-OMs.xlsx")