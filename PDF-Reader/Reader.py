import pytesseract
from pdf2image import convert_from_path
from PIL import Image

path = "S:/TEC/Technical_Function/2.2.Toolshop_Devices/03.Shared/01.Toolshop/03_Inspeção/FERRAMENTARIA/10 - Hallyessa/Nina/3 - Projeto Encerramento LPs/LPs/2024/KW 39/Scanned_from_a_Lexmark_Multifunction_Product31-10-2024-152038.pdf"
pytesseract.pytesseract.tesseract_cmd = r'C:\\Users\\cun1ct\\AppData\\Local\\Programs\\Python\\Python312\\Lib\\site-packages\\pytesseract'

pages = convert_from_path(path, poppler_path="C:/Users/cun1ct/Desktop/Release-24.08.0-0/poppler-24.08.0/Library/bin")
text_data = []

for page in pages:
    text = pytesseract.image_to_string(page)
    text_data.append(text)
print(text_data)