# ----- Libraries -----
import pyautogui as bot
import pandas as pd

# ----- Global Settings -----
bot.FAILSAFE = True
bot.PAUSE = 0.125

# ----- Initial Command -----
bot.click(1804, 16)
bot.sleep(1.125)

# ----- Functions -----
def pressKey(key, count):
    for _ in range(count):
        if key == 'stab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def openTransaction():
    bot.typewrite('CA01')
    pressKey('enter', 1)
    bot.sleep(1.375)

def scriptCreation():
    bot.typewrite('4700.100.981')
    bot.sleep(0.625)
    pressKey('tab', 1)
    bot.sleep(1.125)
    bot.typewrite('6854')
    bot.sleep(0.625)
    pressKey('tab', 1)
    bot.sleep(1.125)
    pressKey('tab', 4)
    bot.sleep(0.275)
    bot.typewrite('MAO10092025')
    bot.sleep(0.375)
    pressKey('tab', 1)
    bot.sleep(0.275)
    bot.typewrite('11.09.2025')
    bot.sleep(0.375)
    pressKey('tab', 2)
    bot.sleep(0.275)
    bot.typewrite('ZBR0001')
    bot.sleep(0.375)

    pressKey('enter', 1)
    bot.sleep(0.875)
    pressKey('enter', 1)
    bot.sleep(0.875)

    pressKey('tab', 7)
    bot.sleep(0.375)
    bot.typewrite('Q55')
    bot.sleep(0.375)
    pressKey('tab', 3)
    bot.sleep(0.375)
    bot.typewrite('1')
    bot.sleep(0.375)
    pressKey('stab', 17)
    bot.sleep(0.375)
    pressKey('enter', 1)

def opInsertion():
    print('Hello World!')

# ----- Start Program -----
if __name__ == '__main__':
    openTransaction()
    scriptCreation()
    opInsertion()