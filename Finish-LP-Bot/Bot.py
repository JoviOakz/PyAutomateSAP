import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.35

bot.click(1802, 14)

# +17
# 1º linha - 233
# 2º linha - 250
# 3º linha - 267
# 4º linha - 284
projectHeight = 250
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
    bot.hotkey('ctrl', 'k')
    bot.click(462, 337)
    bot.hotkey('ctrl', 'v')
    bot.click(538, 479)
    bot.sleep(1)

    try:
        lp_error_exist = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.7))
        if lp_error_exist:
            bot.click(566, 702)
            bot.sleep(0.5)
            bot.click(678, 478)
            bot.click(1231, projectHeight)
            bot.sleep(0.3)
            bot.click(1231, projectHeight)
            bot.hotkey('ctrl', 'l')
            bot.click(1404, projectHeight)
            bot.typewrite('LP nao existe')
            bot.sleep(1)
            return projectHeight + 17, True
    except Exception:
        print('LP existe')

    bot.sleep(1)

    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
        if have_ence:
            bot.click(30, 54)
            bot.sleep(1.5)
            bot.click(1404, projectHeight)
            bot.sleep(0.3)
            bot.click(1404, projectHeight)
            bot.typewrite('PROJETO JA ENCERRADO')
            bot.sleep(1)
            return projectHeight + 17, True
    except Exception:
        print('Projeto não encerrado ainda')

    return projectHeight + 17, False

# muda o status das linhas de compra
def step1_change_status():
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
            bot.hotkey('ctrl', 'l')
            return True
    except Exception:
        return False

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
    bot.sleep(1)

    try:
        have_info = bot.locateOnScreen('images/INFO.png', grayscale=True, confidence=0.8)
        if have_info:
            bot.click(566, 732)
            bot.sleep(0.5)
            bot.click(566, 732)
            bot.sleep(1.25)
    except Exception:
        print('Não houve informação adicional')

# abre a arvore do projeto
def open_tree():
    try:
        have_diagram = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
        if have_diagram:
            bot.click(45, 232)
            bot.sleep(1)
            bot.click(186, 252)
            bot.sleep(1.5)

            try:
                have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
                if have_ence:
                    bot.click(30, 54)
                    bot.sleep(1.5)
                    bot.click(1231, projectHeight - 17)
                    bot.sleep(0.3)
                    bot.click(1231, projectHeight - 17)
                    bot.hotkey('ctrl', 'l')
                    bot.click(1404, projectHeight - 17)
                    bot.typewrite('Diagrama com status ENCE nao pode ser encerrado corretamente')
                    bot.sleep(1)
                    return True
            except Exception:
                try:
                    have_purchase = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
                    if have_purchase:
                        error = step1_change_status()
                        if not error:
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
                            bot.click(186, 250)
                            bot.sleep(1)
                        else:
                            return True
                except Exception:
                    print('Não possui linha de compra!')
    except Exception:
        bot.click(186, 230)
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

# confere se está LIB ou ABER e encerra
def main_function(treeHeight):
    try:
        lib_location = list(bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8))
        if lib_location:
            tec_finish_sequence(first_sequence)
            bot.sleep(1)
            finish_sequence(second_sequence)
            bot.sleep(1)

            try:
                warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                if warning_exist:
                    bot.click(496, 362)
                    bot.sleep(1)
            except Exception:
                print('WARNING não encontrado!')
            treeHeight = adjust_tree_height(treeHeight)
    except Exception:
        try:
            aber_location = list(bot.locateOnScreen('images/ABER.png', grayscale=True, confidence=0.8))
            if aber_location:
                tec_finish_sequence(first_sequence)
                bot.sleep(1)
                finish_sequence(second_sequence)
                bot.sleep(1)

                try:
                    warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                    if warning_exist:
                        bot.click(496, 362)
                        bot.sleep(1)
                except Exception:
                    print('WARNING não encontrado!')
                treeHeight = adjust_tree_height(treeHeight)
        except Exception:
            try:
                ente_location = list(bot.locateOnScreen('images/ENTE.png', grayscale=True, confidence=0.9))
                if ente_location:
                    finish_sequence(second_sequence)
                    bot.sleep(1)

                    try:
                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                        if warning_exist:
                            bot.click(496, 362)
                            bot.sleep(1)
                    except Exception:
                        print('WARNING não encontrado!')
                    treeHeight = adjust_tree_height(treeHeight)
            except Exception:
                treeHeight = adjust_tree_height(treeHeight)
    return treeHeight

# conclui o encerramento
def conclusion():
    bot.click(1404, projectHeight - 17)
    bot.sleep(0.3)
    bot.click(1404, projectHeight - 17)
    bot.typewrite('CONCLUIDO')
    bot.sleep(2.75)

# salvar
def finish_process():
    bot.click(236, 102)
    bot.sleep(0.75)

# programa principal
for _ in range(20):
    treeHeight = 250
    height_adjust_count = 0
    jump_all = False
    jump_main_function = False

    projectHeight, jump_all = open_project(projectHeight)

    if not jump_all:
        jump_main_function = open_tree()

        if not jump_main_function:
            for __ in range(3):
                treeHeight = main_function(treeHeight)

            finish_process()
            conclusion()