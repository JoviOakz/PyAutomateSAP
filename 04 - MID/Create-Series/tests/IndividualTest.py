import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.25

excel_path = "../Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 0

norm = df.at[line, 'Norma']

if norm == 4718301460:
    print('Boa!')
else:
    print('Não deu!')