import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.3

bot.click(1780, 14)

excel_path = "Corpos BICO.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def exists_verification():
    bot.click(1820, 1012)
    press_key('ctrlv', 1)

line = 0
part_number_qty = 764

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    part_number = df.at[line, 'W'] # W -> 764 | X -> 772 | S -> 775 | K -> 29

    exist = exists_verification()

    if not exist:
        bot.click(1820, 1012)
        bot.click(790, 550)
        bot.press('ctrlv', 1)

    line += 1