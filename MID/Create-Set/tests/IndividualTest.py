import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.5

bot.click(1802, 14)

bot.moveTo(582, 462, 0.15)
bot.scroll(-9735)
bot.sleep(0.5)

for ___ in range(15):
            bot.click(500, 520)
            bot.moveTo(582, 462, 0.15)
            bot.scroll(-58)