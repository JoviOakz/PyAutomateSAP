import pyautogui as bot
import pandas as pd
import pyperclip

bot.FAILSAFE = True
bot.PAUSE = 0.55

bot.click(1802, 14)

excel_path = "Data.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

line = 21

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def save():
    bot.click(1850, 946)
    bot.sleep(6)
    bot.click(1136, 664)

for _ in range(2):
    norm = df.at[line, 'Norma']
    qty = df.at[line, 'Quantidade']
    name = df.at[line, 'Nome']

    if norm == 4718301758:
        serie = 3
    else:
        serie = 1

    for __ in range(qty):
        bot.click(1850, 946)

        bot.sleep(1.5)

        bot.click(500, 390)

        pyperclip.copy(name)

        if serie < 10:
            press_key('ctrlv', 1)
            bot.sleep(0.1)
            bot.typewrite('0' + str(serie))
        else:
            press_key('ctrlv', 1)
            bot.sleep(0.1)
            bot.typewrite(str(serie))

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
        bot.moveTo(500, 580, 0.15)

        bot.sleep(0.3)
        bot.scroll(-1000)
        bot.sleep(0.3)

        bot.click()
        
        press_key('tab', 1)
        bot.moveTo(920, 580, 0.15)

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
        bot.sleep(0.3)
        bot.moveTo(582, 462, 0.15)
        bot.sleep(0.3)
        bot.scroll(-9735)
        bot.sleep(0.3)
        bot.click()
        bot.sleep(0.3)
        bot.click(1436, 510)
        bot.sleep(1)

        for ___ in range(15):
            bot.click(500, 520)
            bot.sleep(0.5)
            bot.moveTo(582, 462, 0.15)
            bot.sleep(0.5)
            bot.scroll(-58)
            bot.sleep(0.5)
            bot.click()
            bot.sleep(0.5)
            bot.click(1436, 510)
            bot.sleep(1.1)

        bot.scroll(-6000)
        bot.sleep(0.3)
        bot.click(1400, 766)
        bot.sleep(2)

        save()
        serie += 1
    line += 1