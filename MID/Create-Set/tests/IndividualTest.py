import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.5

bot.click(1802, 14)

bot.moveTo(920, 580)
bot.scroll(-6000)