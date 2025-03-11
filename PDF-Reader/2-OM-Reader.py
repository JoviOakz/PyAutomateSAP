import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import pandas as pd
import re
import os
 
def limpar_string(s):
    return re.sub(r'[^0-9]', '', s)
 
def extract_text_from_pdf_with_ocr(pdf_path):
    extracted_text = []
   
    # Open the PDF
    doc = fitz.open(pdf_path)
   
    for page_num in range(len(doc)):
        try:
            # Get the page
            page = doc[page_num]
           
            # First try to extract text directly (faster if it works)
            text = page.get_text()
           
            if len(text.strip()) <= 100:  # If substantial text was not extracted
                # Get the page as an image for OCR
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
               
                # Use OCR to extract text
                text = pytesseract.image_to_string(img)
           
            # Try to find OMs using regex
            if text:
                # Try different patterns that might match your OM numbers
                patterns = [
                    r'-\s+(\d+)\s+',  # Format: "- 123456 "
                    r'OM[:\s]+(\d+)',  # Format: "OM: 123456" or "OM 123456"
                    r'Order\s+(\d+)',  # Format: "Order 123456"
                    r'#\s*(\d+)'       # Format: "# 123456" or "#123456"
                ]
               
                # Flag to track if a match has been found on this page
                match_found = False
               
                for pattern in patterns:
                    matches = re.findall(pattern, text)
                    if matches and not match_found:
                        # Add only the first match
                        extracted_text.append(matches[0])
                        print(f"Page {page_num+1} - OM: {matches[0]}")
                        match_found = True  # Stop searching after the first match
               
        except Exception as e:
            print(f"Error on page {page_num+1}: {str(e)}")
   
    return extracted_text
 
pdf_path = "OMs - KW05.pdf"
if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
else:
    text = extract_text_from_pdf_with_ocr(pdf_path)
    if text:
        text = [limpar_string(s) for s in text]
        print(f"\nExtracted OM numbers: {text}")
       
        df = pd.DataFrame({"OM": text, "Status": ""})
        df.to_csv("Open-OMs.csv")
        print("Results saved to Open-OMs.xlsx")
    else:
        print("No OMs were extracted from the PDF")