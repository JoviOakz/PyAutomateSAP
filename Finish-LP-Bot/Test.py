import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 1

def tec_finish_sequence(coords):
    for x, y in coords:
        bot.moveTo(x, y, duration=0.1)
        bot.click()

def finish_sequence(coords):
    for x, y in coords:
        bot.moveTo(x, y, duration=0.1)
        bot.click()

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

bot.click(1802, 14)

# Excel elemento PEP // x - fixo // y - +17
# bot.moveTo(1231, 267, duration=0.5)
# bot.moveTo(1231, 284, duration=0.5)
# bot.moveTo(1231, 301, duration=0.5)

# Abrir projeto
bot.click(25, 146, duration=0.5)

# pegar elemento PEP
bot.moveTo(1231, 267, duration=0.5)
bot.click()
bot.click()
bot.hotkey('ctrl', 'c')

# inserir elemento PEP
bot.moveTo(462, 337, duration=0.5)
bot.click()
bot.hotkey('ctrl', 'v')

# confirma
bot.moveTo(538, 479, duration=0.5)
bot.click()

# abre a arvore
bot.moveTo(45, 231, duration=0.5)
bot.click()
bot.sleep(0.75)
bot.moveTo(186, 250, duration=0.5)
bot.click()
bot.sleep(1.5)

# # confere se está LIB
try:
    lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
    if lib_location:
        tec_finish_sequence(first_sequence)
        bot.sleep(2)
        finish_sequence(second_sequence)
        bot.sleep(2)
except Exception:
    bot.moveTo(186, 230, duration=0.5)
    bot.click()
    bot.sleep(1.5)

try:
    lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
    if lib_location:
        tec_finish_sequence(first_sequence)
        bot.sleep(2)
        finish_sequence(second_sequence)
        bot.sleep(2)
except Exception:
    bot.moveTo(186, 210, duration=0.5)
    bot.click()
    bot.sleep(1.5)

try:
    lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
    if lib_location:
        tec_finish_sequence(first_sequence)
        bot.sleep(2)
        finish_sequence(second_sequence)
        bot.sleep(2)
except Exception:
    bot.moveTo(186, 210, duration=0.5)
    bot.click()
    bot.sleep(1.5)

try:
    lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
    if lib_location:
        tec_finish_sequence(first_sequence)
        bot.sleep(2)
        finish_sequence(second_sequence)
        bot.sleep(2)
except Exception:
    bot.moveTo(500, 500, duration=0.5)

bot.moveTo(237, 103, duration=0.1)
bot.click()