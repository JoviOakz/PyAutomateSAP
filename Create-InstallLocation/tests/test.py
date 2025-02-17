import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 0.25

excel_path = "../LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

norm = df.at[0, 'Norma']
pyperclip.copy(norm)

# bot.press('win')
# bot.typewrite('bloco de notas')
# bot.press('enter')

# bot.sleep(1.75)

# bot.click(500, 500)
# bot.typewrite(str(norm) + '-033')

qty = 3

for __ in range(qty - 1):
    if norm == 4718301303:
        print('Sucesso!')