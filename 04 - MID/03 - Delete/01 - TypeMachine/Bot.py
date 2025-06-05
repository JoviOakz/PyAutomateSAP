import pyautogui as bot
import pandas as pd

bot.FAILSAFE = True
bot.PAUSE = 0.3

bot.click(1802, 14)

part_number_qty = 770
line = 0

repeat_count = part_number_qty - line

for _ in range(repeat_count):
    bot.click(812, 830)
    bot.sleep(0.3)

    bot.sleep(1.5)

bot.click(1852, 1074)