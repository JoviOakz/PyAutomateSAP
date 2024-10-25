import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.click(1802, 14)

projectHeight = 267

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

# abre a LP
def open_project(projectHeight):
    bot.click(25, 146)
    bot.click(1231, projectHeight)
    bot.sleep(0.3)
    bot.click(1231, projectHeight)
    bot.hotkey('ctrl', 'c')
    bot.click(462, 337)
    bot.hotkey('ctrl', 'v')
    bot.click(538, 479)
    projectHeight += 17
    return projectHeight

# abre a arvore do projeto
def open_tree():
    bot.moveTo(45, 231, duration=0.15)
    bot.click()
    bot.sleep(0.5)
    bot.moveTo(186, 250, duration=0.15)
    bot.click()
    bot.sleep(1)

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
            bot.sleep(1.5)
            finish_sequence(second_sequence)
            bot.sleep(1.5)

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
    bot.click(236, 102)
    bot.sleep(4)

# programa principal
for _ in range(3):
    treeHeight = 250
    height_adjust_count = 0

    projectHeight = open_project(projectHeight)
    bot.sleep(2)
    open_tree()

    for __ in range(3):
        treeHeight = main_function(treeHeight)

    finish_process()