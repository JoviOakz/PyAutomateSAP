import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 0.45

excel_path = "../LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

lp_value = df.at[0, 'Norma']
pyperclip.copy(lp_value)