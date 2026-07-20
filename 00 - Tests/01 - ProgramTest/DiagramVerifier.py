import pyautogui as bot

bot.PAUSE = 0.15

bot.click(1802, 14)

bot.sleep(1)

line = 0

for _ in range (20):
    line += 1
    bot.hotkey('ctrl', 'c')
    bot.hotkey('ctrl', 'l')
    bot.hotkey('ctrl', 'v')
    bot.press('enter')
    bot.press('esc')
    bot.press('tab')
    bot.typewrite('CADASTRAR')
    bot.press('tab')
    bot.hotkey('ctrl', 'up')

    if line != 20:
        bot.hotkey('ctrl', 'up')
        
    bot.press('del')
    bot.press('down')