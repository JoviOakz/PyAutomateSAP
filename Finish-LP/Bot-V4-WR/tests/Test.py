# --- README ---
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../../' BEFORE THE EXCEL PATH

# LIBRARIES
import pyautogui as bot
import pandas as pd
import pyperclip

# SOFTWARE GLOBAL SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.25

arrowCoords = (15, 166, 400, 200)

first_sequence = [(150, 12), (182, 80), (516, 206), (682, 206)]
second_sequence = [(150, 12), (182, 80), (404, 276), (698, 272)]

coordinates = [
    ((952, 430, 33, 26), (966, 440)),
    ((948, 455, 33, 26), (966, 470)),
    ((605, 455, 33, 26), (1102, 470))
]

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "../../../Open-LPs.xlsx"
df = pd.read_excel(excel_path, engine='openpyxl')

# FUNCTION TO PRESS COMMAND X TIMES
def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        else:
            bot.press(key)

# SEQUENCE TO FINISH WITH ENTE AND ENCE STATUS
def ence_sequence(coords):
    for x, y in coords:
        bot.click(x, y)

# OPEN THE PROJECT
def open_project():
    bot.click(26, 146)

    bot.sleep(1.75)

    lp_value = df.at[line, 'LP']
    pyperclip.copy(lp_value)
    bot.hotkey('ctrl', 'v')
    press_key('enter', 1)

    bot.sleep(1.75)

# IDENTIFIES THE PROJECT STATUS (IF DON'T EXIST, HAVE LBPA OR ALREADY FINISHED) 
def project_status():
    try:
        lp_error_exist = list(bot.locateAllOnScreen('../images/LPNOTEXIST.png', grayscale=True, confidence=0.7))
        
        if lp_error_exist:
            df.at[line, 'Status'] = 'LP não existe!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            bot.sleep(2.5)

            return True
        
    except Exception:
        print('LP exists!')

    try:
        have_ence = bot.locateOnScreen('../images/ENCE.png', grayscale=True, confidence=0.9)
        
        if have_ence:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            press_key('f3', 1)

            bot.sleep(2.5)

            return True
        
    except Exception:
        print('Project don\'t finished yet!')

    try:
        have_lbpa = bot.locateOnScreen('../images/LBPA.png', grayscale=True, confidence=0.9)
        
        if have_lbpa:
            bot.click(150, 15)
            bot.click(210, 75)
            bot.click(470, 75)

            bot.sleep(1.5)

            bot.click(60, 254)
            bot.sleep(0.5)
            bot.click(42, 230)
            bot.moveTo(120, 150, 0.3)
            
            bot.sleep(0.5)

    except Exception:
        print('LBPA status not found!')

    return False

# OPEN PROJECT TREE
def open_tree():
    try:
        have_diagram = list(bot.locateOnScreen('../images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
        
        if have_diagram:
            bot.click(46, 232)
            bot.sleep(2)
            bot.click(186, 252)

            bot.sleep(2)

    except Exception:
        return False

# CHECK IF THE DIAGRAM IS ALREADY ENCE
def diagram_have_ence():
    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
        
        if have_ence:
            try:
                have_purchase = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
                
                if have_purchase:
                    bot.click(150, 15)
                    bot.click(240, 75)
                    bot.click(520, 270)
                    bot.click(680, 300)
                    bot.sleep(2)

            except Exception:
                print('Doesn\'t have any purchase line!')

    except Exception:
        print('Diagram is not ENCE!')

# FINISH THE TREE LINE STATUS
def finish_tree_line():
    try:
        have_aber = list(bot.locateOnScreen('images/ABER.png', grayscale=True, confidence=0.8))
        
        if have_aber:
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
                print('WARNING don\'t exist!')

    except Exception:
        try:
            have_lib = list(bot.locateOnScreen('images/LIB.png', grayscale=True, confidence=0.8))
            
            if have_lib:
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
                        df.to_excel(excel_path, index=False, engine='openpyxl')

                        bot.sleep(5)

                        pending += 1
                
                except Exception:
                    print('Doesn\'t have pending commitment!')

        except Exception:
            try:
                have_ente = list(bot.locateOnScreen('images/ENTE.png', grayscale=True, confidence=0.9))
                
                if have_ente:
                    ence_sequence(second_sequence)

                    bot.sleep(2)

                    try:
                        warning_exist = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                        
                        if warning_exist:
                            press_key('enter', 1)

                            bot.sleep(2)

                    except Exception:
                        print('WARNING don\'t exists!')

            except Exception as e:
                print(f'Error: {e}')

# =======================================================================
# ADICIONAR NO FINISH_TREE_LINE UM VERIFICADOR DE ERRO NO ENCERRAMENTO
# =======================================================================












# =======================================================================
def ence_purchase_line():
    print('x')
# =======================================================================





















# CHANGE THE LP STATUS TO ENCE AND SAVE EVERYTHING
def conclusion():
    bot.hotkey('ctrl', 's')

    df.at[line, 'Status'] = 'Encerrado!'
    df.to_excel(excel_path, index=False, engine='openpyxl')

    bot.sleep(5)

# PUT THE ERROR STATUS OF LP AND DON'T SAVE THE CHANGES
def error_conclusion():
    press_key('f3', 1)

    bot.sleep(2)

    press_key('tab', 1)
    press_key('enter', 1)

    bot.sleep(5)

# EXCEL CONFIG
qty = 25
line = 0

# MAIN PROGRAM
for _ in range(qty):
    jump_all_process = False
    jump_main_function = False
    diagram = True

    pending = 0

    open_project()

    jump_all_process = project_status()

    if not jump_all_process:
        diagram = open_tree()

        if not diagram:
            finish_tree_line()
            jump_main_function = True

        if not jump_main_function:
            diagram_have_ence()
            
            # if ALGUMA COISA: 
                # ence_purchase_line()

            finish_tree_line()

            bot.click(186, 212)
            bot.sleep(2)
            
            finish_tree_line()

            if pending == 0:
                conclusion()
            else:
                error_conclusion()
            
    line += 1