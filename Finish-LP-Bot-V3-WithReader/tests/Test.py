import pyautogui as bot
import pandas as pd
import pyperclip

# TESTS NEED '../' BEFORE THE PATH OF IMAGES

bot.FAILSAFE = True
bot.PAUSE = 0.25

# TESTS NEED '../../' BEFORE THE PATH OF PDF

pdf_path = "../../Open-LPs.xlsx"
df = pd.read_excel(pdf_path, engine='openpyxl')

lp_value = df.at[0, 'LP']
pyperclip.copy(lp_value)