import pyautogui as bot
import pandas as pd
from openpyxl.styles import PatternFill

bot.FAILSAFE = True
bot.PAUSE = 0.25

excel_path = "../ApontamentoYesica.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

df.at[11, 'Elemento PEP'] = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")

# bot.typewrite('Planejadora Yesica - 20/02/2025')
# bot.typewrite('H')
# bot.typewrite('100')
# bot.typewrite('025PROJ')

# bot.typewrite('92886895')