import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.3

bot.click(1802, 14)

excel_path = "../Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 0

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def save():
    bot.click(1850, 946)
    bot.sleep(3)
    bot.click(1136, 664)

for _ in range(23):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']
    name = df.at[line, 'Nome']

    if norm == 4718301460:
        serie = 2
    elif norm == 4718301303:
        serie = 3
    else:
        serie = 1

    for __ in range(qty):
        bot.click(1850, 946)

        bot.sleep(1.5)

        bot.click(500, 390)

        if serie < 10:
            bot.typewrite(str(name) + '0' + str(serie))
        else:
            bot.typewrite(str(name) + str(serie))

        press_key('tab', 2)
        bot.typewrite(str(norm))
        bot.click(500, 520)
        press_key('tab', 1)

        if serie < 10:
            bot.typewrite('00' + str(serie))
        else:
            bot.typewrite('0' + str(serie))
        
        press_key('tab', 1)
        bot.typewrite('100000')
        press_key('tab', 1)
        bot.typewrite('15')
        press_key('tab', 1)
        bot.moveTo(500, 580, 0.5)

        bot.sleep(0.3)
        bot.scroll(-1000)
        bot.sleep(0.3)

        bot.click()
        
        press_key('tab', 1)
        bot.moveTo(920, 580, 0.5)

        bot.sleep(0.3)
        bot.scroll(-6000)
        bot.sleep(0.3)

        bot.click()
        
        press_key('tab', 1)
        bot.typewrite('12')

        bot.sleep(0.3)
        
        bot.click(1272, 724)
        press_key('tab', 2)

        if serie < 10:
            bot.typewrite('00' + str(serie))
        else:
            bot.typewrite('0' + str(serie))
        
        bot.click(510, 950)
        bot.sleep(0.3)
        bot.click(1436, 1006)
        bot.sleep(0.3)
        bot.scroll(-6000)
        bot.sleep(0.3)

        bot.click(500, 520)
        bot.moveTo(582, 462, 0.15)
        bot.scroll(-9735)
        bot.click()
        bot.sleep(0.3)
        bot.click(1436, 510)
        bot.sleep(0.3)

        for ___ in range(15):
            bot.click(500, 520)
            bot.moveTo(582, 462, 0.15)
            bot.scroll(-58)
            bot.click()
            bot.sleep(0.3)
            bot.click(1436, 510)
            bot.sleep(0.3)

        bot.scroll(-6000)
        bot.sleep(0.3)
        bot.click(1400, 766)
        bot.sleep(2)

        save()
        serie += 1