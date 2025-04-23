import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.25

excel_path = "../LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 1

qty = df.at[line, 'Quantidade']

print(qty)