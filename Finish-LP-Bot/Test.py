import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

projectHeight = 267

def step1_change_status():
    bot.click(600, 884)
    bot.sleep(1.25)

    try:
        error_exist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7))
        if error_exist:
            print('Existe')
            bot.click(456, 390)
            bot.sleep(1)
            bot.click(566, 700)
            bot.sleep(1)
            bot.click(30, 54)
            bot.sleep(1.5)
            bot.click(1231, projectHeight)
            bot.sleep(0.3)
            bot.click(1231, projectHeight)
            bot.click(1156, 130)
            bot.sleep(0.3)
            bot.click(1154, 314)
            bot.sleep(0.3)
            bot.click(1231, projectHeight + 17)
            bot.click(1156, 130)
            bot.sleep(0.3)
            bot.click(1222, 316)
    except Exception as e:
        print(f'Error: {e}')

bot.click(1804, 15)
step1_change_status()