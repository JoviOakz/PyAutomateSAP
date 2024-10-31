import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.click(1802, 14)

# +17
projectHeight = 267
arrowCoords = (15, 166, 400, 200)

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

coordinates = [
    ((474, 430, 33, 26), (488, 440)),
    ((470, 455, 33, 26), (488, 470)),
    ((605, 455, 33, 26), (622, 470))
]

# abre a LP
def open_project(projectHeight):
    bot.click(25, 146)
    bot.click(1231, projectHeight)
    bot.sleep(0.3)
    bot.click(1231, projectHeight)
    bot.hotkey('ctrl', 'c')
    bot.click(462, 337)
    bot.hotkey('ctrl', 'v')
    bot.click(1136, 126)
    bot.sleep(0.3)
    bot.click(1136, 126)
    bot.click(538, 479)
    return projectHeight + 17

# muda o status das linhas de compra
def step1_change_status():
    bot.click(186, 252)
    bot.sleep(1.25)
    bot.click(580, 236)
    bot.sleep(1.25)
    bot.click(486, 884)
    bot.sleep(1.25)
    bot.click(600, 884)
    bot.sleep(1.5)

    try:
        error_exist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7))
        if error_exist:
            bot.click(456, 390)
            bot.sleep(1)
            bot.click(566, 700)
            bot.sleep(1)
            bot.click(30, 54)
            bot.sleep(1.5)
            bot.click(1231, projectHeight - 17)
            bot.sleep(0.3)
            bot.click(1231, projectHeight - 17)
            bot.click(1156, 130)
            bot.sleep(0.3)
            bot.click(1154, 314)
            bot.sleep(0.3)
            bot.click(1231, projectHeight)
            bot.click(1156, 130)
            bot.sleep(0.3)
            bot.click(1222, 316)
            return False
    except Exception as e:
        return True

# muda o status das linhas de compra
def step2_change_status():
    bot.click(290, 332)
    bot.typewrite('92903610')
    
    for region, click_position in coordinates:
        try:
            if bot.locateOnScreen('images/CHECK.png', grayscale=True, confidence=0.7, region=region):
                print('encontrado')
            else:
                raise Exception
        except Exception:
            bot.click(click_position)

    bot.sleep(0.5)
    bot.click(646, 988)
    bot.sleep(1.25)
    bot.click(468, 732)
    bot.sleep(1.25)

# abre a arvore do projeto
def open_tree():
    bot.moveTo(45, 232, duration=0.15)
    bot.click()
    bot.sleep(1)

    try:
        have_purchase = bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords)
        if have_purchase:
            error = step1_change_status()
            if error:
                warning_exist = None
                while not warning_exist:
                    try:
                        step2_change_status()
                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                    except Exception as e:
                        print(f'Erro: {e}')

                bot.click(492, 308)
                bot.sleep(1)
                bot.click(566, 702)
                bot.sleep(1)
                bot.click(582, 208)
                bot.sleep(1)
            else:
                return False
    except Exception:
        print('Não possui linha de compra!')

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
for _ in range(20):
    treeHeight = 250
    height_adjust_count = 0
    jump = True

    projectHeight = open_project(projectHeight)
    bot.sleep(2)
    jump = open_tree()

    if jump:
        for __ in range(3):
            treeHeight = main_function(treeHeight)

    finish_process()