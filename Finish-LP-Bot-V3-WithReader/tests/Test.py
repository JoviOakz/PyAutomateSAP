# --- README --- 
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../' BEFORE THE PATH OF PDF

# LIBRARIES
import pyautogui as bot
import pandas as pd
import pyperclip

# SOFTWARE GLOBAL SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.25

arrowCoords = (15, 166, 400, 200)

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

coordinates = [
    ((474, 430, 33, 26), (488, 440)),
    ((470, 455, 33, 26), (488, 470)),
    ((605, 455, 33, 26), (622, 470))
]

# LOWER DE CODE SCREEN
bot.click(1802, 14)

# PDF CONFIGURATION
pdf_path = "../../Open-LPs.xlsx"
df = pd.read_excel(pdf_path, engine='openpyxl')

# OPEN THE PROJECT AND IDENTIFIES IF EXISTS OR IS ALREADY FINISHED
def open_project():
    bot.click(26, 146)

    bot.sleep(2)

    lp_value = df.at[0, 'LP']
    pyperclip.copy(lp_value)
    bot.hotkey('ctrl', 'v')
    bot.press('enter')

    bot.sleep(2)

    try:
        lp_error_exist = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.7))
        
        if lp_error_exist:
            bot.typewrite('LP don\'t exists')

            return True
        
    except Exception:
        print('LP exists')

    bot.sleep(2)

    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
        
        if have_ence:
            bot.typewrite('Project already finished!')

            return True
        
    except Exception:
        print('Project don\'t finished yet!')

    return False

# OPEN TREE AND ALSO CHANGE THE PURCHASE STATUS IF EXISTS
def open_tree():
    try:
        have_diagram = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
        
        if have_diagram:
            bot.click(46, 232)
            bot.sleep(2)
            bot.click(186, 252)
            bot.sleep(2)

            try:
                have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
                
                if have_ence:
                    try:
                        have_purchase = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
                        
                        if have_purchase:
                            bot.click(150, 15)
                            bot.sleep(2)
                            bot.click(240, 75)
                            bot.sleep(2)
                            bot.click(520, 270)
                            bot.sleep(2)
                            bot.click(680, 300)
                            bot.sleep(2)

                            try:
                                ente_location = list(bot.locateOnScreen('images/ENTE.png', grayscale=True, confidence=0.9))
                                
                                if ente_location:
                                    bot.click(150, 15)
                                    bot.sleep(2)
                                    bot.click(240, 75)
                                    bot.sleep(2)
                                    bot.click(520, 200)
                                    bot.sleep(2)
                                    bot.click(680, 240)
                                    bot.sleep(2)

                            except Exception:
                                print('Don\'t have ENTE status')

                            bot.click(580, 240)
                            bot.sleep(2)
                            bot.click(484, 884)
                            bot.sleep(2)

                            bot.moveTo(606, 848)
                            bot.mouseDown()
                            bot.moveTo(846, 848, duration=0.25)
                            bot.mouseUp()
                            bot.sleep(2)

                            try:
                                check_box = list(bot.locateAllOnScreen('images/CHECK.png', grayscale=True, confidence=0.8))
                                
                                if check_box:
                                    try:
                                        have_baixa = list(bot.locateAllOnScreen('images/BAIXACONF.png', grayscale=True, confidence=0.8))
                                        
                                        if have_baixa:
                                            if len(check_box) != len(have_baixa):
                                                bot.click(150, 15)
                                                bot.sleep(2)
                                                bot.click(240, 75)
                                                bot.sleep(2)
                                                bot.click(520, 270)
                                                bot.sleep(2)
                                                bot.click(680, 300)
                                                bot.sleep(2)
                                                bot.click(484, 884)
                                                bot.sleep(2)
                                                bot.click(600, 884)
                                                bot.sleep(2)
                                                bot.click(492, 360)
                                                bot.sleep(2)
                                                bot.click(492, 360)
                                                bot.sleep(2)

                                                warning_exist = None
                                                
                                                while not warning_exist:
                                                    try:
                                                        change_status_step_2()

                                                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                                                    
                                                    except Exception as e:
                                                        print(f'Error: {e}')

                                                # REVER O QUE É ESSA BOMBA AQUI!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                                bot.click(492, 308)
                                                bot.sleep(2)
                                                bot.click(566, 702)
                                                bot.sleep(2)
                                                # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                                    except Exception as e:
                                        print(f'Erro: {e}')

                            except Exception as e:
                                print(f'Error: {e}')

                    except Exception:
                        print('Doesn\'t have any purchase!')

            except Exception:
                try:
                    have_purchase = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
                    
                    if have_purchase:
                        error = change_status_step_1()
                        
                        if not error:
                            warning_exist = None
                            
                            while not warning_exist:
                                try:
                                    change_status_step_2()

                                    warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                                
                                except Exception as e:
                                    print(f'Error: {e}')

                            bot.click(492, 308)
                            bot.sleep(2)
                            bot.click(566, 702)
                            bot.sleep(2)
                            
                        else:
                            return True
                        
                except Exception:
                    print('Doesn\'t have any purchase!')

                bot.click(582, 208)
                bot.sleep(2)
                bot.click(186, 252)
                bot.sleep(2)

    except Exception:
        print('Don\'t have diagram!')

# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# 
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

# SAVE ALL THE CHANGES
def finish_process():
    bot.click(236, 102)
    bot.sleep(2)

# CONCLUDES THE STATUS OFF LP
def conclusion():
    print('')

# MAIN PROGRAM
for _ in range(10):
    treeHeight = 250
    jump_main_function = False
    jump_all = False

    jump_all = open_project()
    bot.sleep(2)

    if not jump_all:
        jump_main_function = open_tree()

        if not jump_main_function:
            for __ in range(2):
                treeHeight = main_function(treeHeight)

            finish_process()
            conclusion()