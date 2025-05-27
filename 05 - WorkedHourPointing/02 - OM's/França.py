import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 1

bot.click(1802, 14)

excel_path = 'ApontamentoFrança.xlsx'
df = pd.read_excel(excel_path, engine='openpyxl')

group = []
line = 0

rwork = df.at[line, 'Trab Real']
uni = df.at[line, 'Un']
work_center = df.at[line, 'CenTrab']
description = df.at[line, 'Descrição']
center = df.at[line, 'Centro']
tativ = df.at[line, 'Tativ']
date = df.at[line, 'Data']
edv = df.at[line, 'Nº Pessoal']

for _ in range(15):
    order = df.at[line, 'Ordem']
    group.append(order)
    line += 1

all_values = '\n'.join(map(str, group))

bot.press('tab')
pyperclip.copy(all_values)
bot.hotkey('ctrl', 'v')