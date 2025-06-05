import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 1

image = bot.screenshot(region=(600, 700, 200, 200))
x, y = bot.locateCenterOnScreen(image)
print(x, y)

bot.moveTo(x, y)