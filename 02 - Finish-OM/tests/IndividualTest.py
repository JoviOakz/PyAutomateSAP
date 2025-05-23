import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.sleep(1)

try:
    warning = list(bot.locateAllOnScreen('../images/WARNING.png', grayscale=True, confidence=0.8))
    conforder = list(bot.locateAllOnScreen('../images/CONFORDER.png', grayscale=True, confidence=0.8))

    if warning or conforder:
        print('Existe!')

except Exception as e:
    print(f'Error: {e}')