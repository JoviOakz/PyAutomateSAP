import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.3

bot.click(1780, 14)

excel_path = "Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 20

def enter_material(norm):
    bot.click(200, 488)
    bot.sleep(0.3)
    bot.typewrite(str(norm))
    bot.sleep(0.3)
    bot.click(1780, 480)
    bot.sleep(0.3)
    bot.click(1650, 656)
    bot.sleep(1)
    bot.click(340, 696)
    bot.sleep(0.3)

def main_function(serie):
    if serie < 10:
        bot.typewrite('00' + str(serie))
    else:
        bot.typewrite('0' + str(serie))
    
    press_key('tab', 2)
    bot.typewrite('60')
    press_key('tab', 1)
    bot.typewrite('2.52')
    press_key('enter', 1)
    bot.sleep(0.15)
    bot.click(340, 696)

def save():
    bot.click(1832, 896)
    bot.sleep(1.5)
    press_key('enter', 1)
    bot.sleep(0.5)

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

line = 0
serie_qty = 1

repeat_count = serie_qty - line

for _ in range(repeat_count):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']

    if norm == 4718301460:
        serie = 4
    else:
        serie = 2

    enter_material(norm)

    for __ in range(qty):
        main_function(serie)
        serie += 1

    line += 1
    save()

bot.alert(title='BotText', text='Programa encerrado!')