import pyautogui as bot

bot.FAILSAFE = TRUE
# bot.sleep(1)

bot.click(1880, 18)

# EXCEL CONFIGURATION
# -------------------

# texto
def press_key(key):
  if key == 'ctrlv':
    bot.hotkey('ctrl', 'v')
  else:
    bot.press(key)

# VERIFY IF THE XXXX
def exists_verification():
  bot.click(1820, 1012)
  bot.press_key('ctrlv', 1)

# EXCEL CONFIGS
line = 0
part_number_qty = 900
repeat_count = part_number_qty - line

# MAIN FUNCTION
for _ in range(repeat_count):
  # part_number = bot.copy['X, Y, Z', 'line']

  exist = exists_verification()

  if not exist:
    bot.click(1820, 1012)
    bot.click(790, 550)
    bot.press(ctrlv, 1)

  line += 1
