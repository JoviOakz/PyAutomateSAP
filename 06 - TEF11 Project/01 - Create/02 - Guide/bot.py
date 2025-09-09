# ----- Libraries -----
import pyautogui as bot
import pandas as pd

# ----- Global Settings -----
bot.FAILSAFE = True
bot.PAUSE = 0.25

# ----- Initial Command -----
bot.click(1804, 16)
bot.sleep(1)

# ----- Functions -----
def pressKey(key, count):
    for _ in range(count):
        if key == 'crtlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

def insertInfo():
    bot.keyDown('ctrl')
    bot.click(322, 180)
    bot.keyUp('ctrl')
    bot.sleep(1)

    bot.typewrite('4700.100.980')
    bot.sleep(0.25)
    pressKey('tab', 1)
    bot.sleep(1)

    bot.typewrite('6854')
    bot.sleep(0.25)
    pressKey('tab', 5)
    bot.sleep(0.15)
    
    bot.typewrite('MAO09092025')
    bot.sleep(0.25)
    pressKey('tab', 1)
    bot.sleep(0.15)
    bot.typewrite('10.09.2025')
    bot.sleep(0.25)
    pressKey('tab', 2)
    bot.sleep(0.15)

    bot.typewrite('ZBR0001')
    bot.sleep(0.25)
    # pressKey('enter')

# ----- Start Program -----
if __name__ == '__main__':
    insertInfo()