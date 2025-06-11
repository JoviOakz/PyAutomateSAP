import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.35

bot.click(1802, 14)

excel_path = "LocInst.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

line = 2

for _ in range(4):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']

    if norm == 4729106784:
        serie = 53
    else:
        serie = 1

    for __ in range(qty):
        if serie < 10:
            bot.typewrite(str(norm) + '-00' + str(serie))
        else:
            bot.typewrite(str(norm) + '-0' + str(serie))

        press_key('enter', 1)

        bot.sleep(2.25)

        press_key('tab', 3)
        press_key('right', 1)
        press_key('enter', 1)

        bot.sleep(2)

        press_key('tab', 1)
        bot.typewrite('6854')
        bot.sleep(0.75)
        press_key('tab', 1)
        bot.typewrite('CT/303')
        press_key('tab', 2)
        bot.typewrite('ZBP')

        press_key('shtab', 4)
        press_key('right', 1)
        press_key('enter', 1)

        bot.sleep(2)

        press_key('tab', 4)

        if norm == 4729108508 or norm == 4328700313:
            bot.typewrite('685438')
        else:
            bot.typewrite('685434')
        
        bot.sleep(0.75)
        press_key('tab', 6)
        bot.typewrite('FF0600')
        bot.sleep(0.75)
        press_key('tab', 1)
        bot.typewrite('6854')
        bot.sleep(0.75)
        press_key('tab', 1)
        bot.typewrite('ZBR000008')

        bot.sleep(1.25)

        press_key('enter', 1)
        bot.sleep(2.25)
        press_key('ctrls', 1)

        bot.sleep(3)

        serie += 1
    line += 1

bot.alert(title='BotText', text='Programa encerrado!')