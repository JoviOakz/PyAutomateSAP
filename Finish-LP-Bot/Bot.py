import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.click(1802, 14)

projectHeight = 250

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

def open_project():
    bot.click(25, 146, duration=0.2)
    bot.moveTo(1231, 267, duration=0.2)
    bot.click()
    bot.sleep(0.75)
    bot.click()
    bot.hotkey('ctrl', 'c')
    bot.moveTo(462, 337, duration=0.2)
    bot.click()
    bot.hotkey('ctrl', 'v')
    bot.moveTo(538, 479, duration=0.2)
    bot.click()
    bot.sleep(1.5)

def open_tree():
    bot.moveTo(45, 231, duration=0.2)
    bot.click()
    bot.sleep(0.75)
    bot.moveTo(186, 250, duration=0.2)
    bot.click()
    bot.sleep(2)

# encerra tecnicamente
def tec_finish_sequence(coords):
    for x, y in coords:
        bot.moveTo(x, y, duration=0.1)
        bot.click()

# encerra a atividade
def finish_sequence(coords):
    for x, y in coords:
        bot.moveTo(x, y, duration=0.1)
        bot.click()

# confere se está LIB e encerra
def main_function(projectHeight):
    try:
        lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
        if lib_location:
            tec_finish_sequence(first_sequence)
            bot.sleep(1.5)
            finish_sequence(second_sequence)
            bot.sleep(1.5)
    except Exception:
        projectHeight -= 20
        bot.moveTo(186, projectHeight, duration=0.1)
        bot.click()
        bot.sleep(1.5)

# salvar
def finish_process():
    bot.moveTo(237, 103, duration=0.1)
    bot.click()
    bot.sleep(1.5)

for i in range(5):
    open_project()
    open_tree()

    for j in range(3):
        main_function(projectHeight)

    finish_process()