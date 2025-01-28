import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 0.25

pdf_path = "../Open-LPs.xlsx"
df = pd.read_excel(pdf_path, engine='openpyxl')

lp_value = df.at[0, 'LP']
pyperclip.copy(lp_value)

bot.press('win')
bot.typewrite('bloco de notas')
bot.press('enter')
bot.sleep(0.3)
bot.moveTo(x=500, y=500, duration=0.5)
bot.click()
bot.hotkey('ctrl', 'v')
bot.press('enter')
bot.write('deu certo')