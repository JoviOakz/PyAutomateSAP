import pyautogui as bot

bot.PAUSE = 0.25

bot.click(400, 400)
bot.press('enter')
bot.sleep(0.5)
bot.hotkey('ctrl', 'shift', 'f8')
bot.sleep(0.5)
bot.press('f5')
bot.sleep(0.5)
bot.press('enter')
bot.sleep(3)
bot.hotkey('ctrl', 's')