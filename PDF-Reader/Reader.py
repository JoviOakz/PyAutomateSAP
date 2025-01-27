
import pandas as pd
from pdf2image import convert_from_path
import pytesseract

def extract_text_from_pdf(pdf_path):
    images = convert_from_path(pdf_path)
 
    extracted_text = []
    for i, image in enumerate(images):
        try:
            rotated_image = image.rotate(90, expand=True)
 
            text = pytesseract.image_to_string(rotated_image)
    
            index0 = text.index("LP-")
            index1 = text.index("\n", index0+1)
    
            lp = text[index0:index1]
            extracted_text.append(lp.strip())
        except:
            print("Error on page ", i+1)

    return extracted_text
 
pdf_path = "PDF-Reader/LPs.pdf"
text = extract_text_from_pdf(pdf_path)
print(text)

df = pd.DataFrame({"LP": text, "Status": ""})

df.to_excel("Open LPs.xlsx")