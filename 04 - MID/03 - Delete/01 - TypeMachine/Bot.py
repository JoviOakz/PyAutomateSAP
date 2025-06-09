import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.1

bot.click(1802, 14)

bot.sleep(0.5)

part_number_qty = 10
line = 0

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    bot.click(710, 1046)
    bot.press('tab')
    bot.press('enter')