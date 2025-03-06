# --- README ---
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../../' BEFORE THE PATH OF PDF

# LIBRARIES
import pyautogui as bot
import pandas as pd
import pyperclip

# SOFTWARE GLOBAL SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.35

arrowCoords = (15, 166, 400, 200)

first_sequence = [(150, 13), (183, 79), (515, 207), (683, 207)]
second_sequence = [(150, 13), (183, 79), (404, 276), (698, 271)]

coordinates = [
    ((474, 430, 33, 26), (488, 440)),
    ((470, 455, 33, 26), (488, 470)),
    ((605, 455, 33, 26), (622, 470))
]

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "../../Open-LPs.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

# OPEN THE PROJECT AND IDENTIFIES IF EXISTS OR IS ALREADY FINISHED
def open_project():
    bot.click(26, 146)

    bot.sleep(2)

    lp_value = df.at[line, 'LP']
    pyperclip.copy(lp_value)
    bot.hotkey('ctrl', 'v')
    press_key('enter', 1)

    bot.sleep(2)

    try:
        lp_error_exist = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.7))
        
        if lp_error_exist:
            df.at[line, 'Status'] = 'LP não existe!'

            return True
        
    except Exception:
        print('LP exists!')

    bot.sleep(2)

    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
        
        if have_ence:
            df.at[line, 'Status'] = 'Encerrado'

            return True
        
    except Exception:
        print('Project don\'t finished yet!')

    return False

# CHANGE THE PURCHASE LINE STATUS
def change_status_step_one():
    bot.click(580, 236)

    bot.sleep(2)

    bot.click(486, 884)

    bot.sleep(2)

    bot.moveTo(606, 848)
    bot.mouseDown()
    bot.moveTo(846, 848, duration=0.25)
    bot.mouseUp()

    bot.sleep(2)

    try:
        have_baixa = list(bot.locateAllOnScreen('images/BAIXACONF.png', grayscale=True, confidence=0.8))
        
        if have_baixa:
            bot.click(150, 15)

            bot.sleep(2)

            bot.click(240, 75)

            bot.sleep(2)

            bot.click(500, 270)

            bot.sleep(2)

            bot.click(690, 300)

            bot.sleep(2)

    except Exception as e:
        print(f'Erro: {e}')

    bot.click(486, 884)

    bot.sleep(2)

    bot.click(600, 884)

    bot.sleep(2)

    try:
        error_exist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7))
        
        if error_exist:
            press_key('tab', 2)
            press_key('enter', 1)

            bot.sleep(1)

            bot.click(566, 700)

            bot.sleep(1)

            press_key('f3', 1)

            bot.sleep(1.5)

            try:
                save_exist = list(bot.locateAllOnScreen('images/GRAVAR.png', grayscale=True, confidence=0.7))

                if save_exist:
                    press_key('tab', 1)

                    press_key('enter', 1)
                    
                    bot.sleep(2)

            except Exception:
                print('Doesn\'t have any changes!')

            df.at[line, 'Status'] = 'Problema na linha de compra!'

            return True
        
    except Exception:
        return False
    
# FUNCTION TO PRESS COMMAND X TIMES
def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)
    
# CHANGE THE PURCHASE LINE STATUS
def change_status_step_two():
    press_key('tab', 2)

    bot.typewrite('92903610')
    
    for region, click_position in coordinates:
        try:
            if bot.locateOnScreen('images/CHECK.png', grayscale=True, confidence=0.7, region=region):
                print('Found!')

            else:
                raise Exception
            
        except Exception:
            bot.click(click_position)

    bot.sleep(2)

    bot.click(1128, 988)

    bot.sleep(2)

    press_key('tab', 1)
    press_key('enter', 1)

    bot.sleep(2)

    try:
        have_info = bot.locateOnScreen('images/INFO.png', grayscale=True, confidence=0.8)
        
        if have_info:
            press_key('tab', 1)
            press_key('enter', 1)

            bot.sleep(2)

            press_key('tab', 1)
            press_key('enter', 1)

            bot.sleep(2)

    except Exception:
        print('Don\'t have any additional information!')

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
                                print('Don\'t have ENTE status!')

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

                                                press_key('enter', 1)

                                                bot.sleep(2)

                                                press_key('enter', 1)

                                                bot.sleep(2)

                                                warning_exist = None
                                                
                                                while not warning_exist:
                                                    try:
                                                        change_status_step_two()

                                                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                                                    
                                                    except Exception as e:
                                                        print(f'Error: {e}')

                                                press_key('tab', 2)
                                                press_key('enter', 1)

                                                bot.sleep(2)

                                                press_key('tab', 1)
                                                press_key('enter', 1)

                                                bot.sleep(2)

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
                        error = change_status_step_one()
                        
                        if not error:
                            warning_exist = None
                            
                            while not warning_exist:
                                try:
                                    change_status_step_two()

                                    warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                                
                                except Exception as e:
                                    print(f'Error: {e}')

                            press_key('tab', 2)
                            press_key('enter', 1)

                            bot.sleep(2)

                            press_key('tab', 1)
                            press_key('enter', 1)

                            bot.sleep(2)
                            
                        else:
                            return True
                        
                except Exception:
                    print('Doesn\'t have any purchase!')

            bot.click(582, 208)

            bot.sleep(2)

    except Exception:
        print('Don\'t have diagram!')

# FINISH WITH ENTE AND ENCE STATUS
def ence_sequence(coords):
    for x, y in coords:
        bot.click(x, y)

# ADJUSTS THE TREE HEIGHT
def adjust_tree_height(height):
    global height_adjust_count
    
    if height_adjust_count < 1:
        height -= 40
        bot.click(186, height)

        bot.sleep(2)

        height_adjust_count += 1

    return height

def main_function(treeHeight, pending):
    try:
        lib_location = list(bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8))
        
        if lib_location:
            ence_sequence(first_sequence)

            bot.sleep(2)

            ence_sequence(second_sequence)

            bot.sleep(2)

            try:
                warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                
                if warning_exist:
                    press_key('enter', 1)

                    bot.sleep(2)

            except Exception:
                print('WARNING don\'t exists!')

            try:
                error_exist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7))

                if error_exist:
                    press_key('tab', 1)
                    press_key('enter', 1)

                    df.at[line, 'Status'] = 'Compromisso pendente!'

                    pending += 1
            
            except Exception:
                print('Doesn\'t have pending commitment!')

            treeHeight = adjust_tree_height(treeHeight)

    except Exception:
        try:
            aber_location = list(bot.locateOnScreen('images/ABER.png', grayscale=True, confidence=0.8))
            
            if aber_location:
                ence_sequence(first_sequence)

                bot.sleep(2)

                ence_sequence(second_sequence)

                bot.sleep(2)

                try:
                    warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                    
                    if warning_exist:
                        press_key('enter', 1)

                        bot.sleep(2)

                except Exception:
                    print('WARNING don\' exists!')

                treeHeight = adjust_tree_height(treeHeight)

        except Exception:
            try:
                ente_location = list(bot.locateOnScreen('images/ENTE.png', grayscale=True, confidence=0.9))
                
                if ente_location:
                    ence_sequence(second_sequence)

                    bot.sleep(2)

                    try:
                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                        
                        if warning_exist:
                            press_key('enter', 1)

                            bot.sleep(2)

                    except Exception:
                        print('WARNING don\' exists!')

                    treeHeight = adjust_tree_height(treeHeight)

            except Exception:
                treeHeight = adjust_tree_height(treeHeight)

    return treeHeight, pending

# SAVE ALL CHANGES AND CONCLUDES THE STATUS OF LP
def conclusion():
    bot.hotkey('ctrl', 's')

    bot.sleep(2)

    df.at[line, 'Status'] = 'Encerrado'

# CONCLUDES THE ERROR STATUS OF LP
def error_conclusion():
    press_key('f3', 1)

    bot.sleep(2)

    press_key('tab', 1)
    press_key('enter', 1)

# MAIN PROGRAM
for _ in range(10):
    treeHeight = 250
    jump_main_function = False
    jump_all = False
    pending = 0
    line = 0

    jump_all = open_project()

    bot.sleep(2)

    if not jump_all:
        jump_main_function = open_tree()

        if not jump_main_function:
            for __ in range(2):
                treeHeight, pending = main_function(treeHeight, pending)

                bot.sleep(2)

            if pending == 0:
                conclusion()
            else:
                error_conclusion()
            
    line += 1