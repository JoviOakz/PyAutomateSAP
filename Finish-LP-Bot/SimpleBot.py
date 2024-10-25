import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

projectHeight = 267

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

# encerra tecnicamente
def tec_finish_sequence(coords):
    for x, y in coords:
        bot.click(x, y)

# encerra a atividade
def finish_sequence(coords):
    for x, y in coords:
        bot.click(x, y)

# ajusta a altura e entra na árvore
def adjust_tree_height(height):
    global height_adjust_count
    if height_adjust_count < 2:
        height -= 20
        bot.click(186, height)
        bot.sleep(1)
        height_adjust_count += 1
        return height

# confere se está LIB e encerra
def main_function(treeHeight):
    try:
        lib_location = bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8)
        if lib_location:
            tec_finish_sequence(first_sequence)
            bot.sleep(1.25)
            finish_sequence(second_sequence)
            bot.sleep(1.25)

            try:
                warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                if warning_exist:
                    print('Warning encontrado!')
                    bot.click(496, 362)
                    bot.sleep(1)
            except Exception:
                print('WARNING não encontrado!')
            treeHeight = adjust_tree_height(treeHeight)
    except Exception:
        treeHeight = adjust_tree_height(treeHeight)
    return treeHeight

# salvar
def finish_process():
    bot.click(237, 103)

# programa principal
treeHeight = 250
height_adjust_count = 0

for __ in range(3):
    treeHeight = main_function(treeHeight)

finish_process()