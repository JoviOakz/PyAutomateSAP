import pyautogui as bot

# bot.hotkey('alt', 'tab')

bot.moveTo(780, 1050, duration=0.5)
# bot.click(x=780, y=1050)
bot.click()

bot.sleep(1)

bot.write('bloco de notas')
bot.press('enter')

bot.sleep(1)

bot.moveTo(400, 300, duration=0.5)
bot.click()
bot.write('Bom dia gatinho!')
bot.moveTo(600, 500, duration=0.5)

# bot.click(x=1089, y=1052)
# bot.sleep(5)
# bot.click(x=611, y=171)
# bot.write('ps0')
# bot.press('enter')