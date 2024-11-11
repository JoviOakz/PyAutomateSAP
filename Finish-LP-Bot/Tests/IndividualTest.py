import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.35

bot.moveTo(1231, 600 - 17)
bot.sleep(2)
bot.moveTo(1231, 600 - 17)
bot.hotkey('ctrl', 'l')