# --- README ---
# TESTS NEED '../' BEFORE THE PATH OF IMAGES
# TESTS NEED '../../../' BEFORE THE EXCEL PATH

# LIBRARIES
import pyautogui as bot
import pandas as pd
import pyperclip

# GLOBAL SOFTWARE SETTINGS
bot.FAILSAFE = True
bot.PAUSE = 0.75

arrowCoords = (15, 166, 400, 200)
hourCoords = (880, 332, 50, 188)
workCenterCoords = (986, 332, 108, 188)

first_sequence = [(150, 12), (182, 80), (516, 206), (682, 206)]
second_sequence = [(150, 12), (182, 80), (404, 276), (698, 272)]

coordinates = [
    ((952, 544, 33, 26), (966, 552)),
    ((948, 570, 33, 26), (966, 582)),
    ((605, 570, 33, 26), (1102, 582))
]

# PUT DOWN THE CODE SCREEN
bot.click(1802, 14)

# EXCEL CONFIGURATION
excel_path = "../../Open-LPs.xlsx"
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

    bot.sleep(2)

# IDENTIFIES THE PROJECT STATUS (IF DON'T EXIST, HAVE LBPA OR ALREADY FINISHED) 
def project_status():
    try:
        lp_error_exist = list(bot.locateAllOnScreen('images/LPNOTEXIST.png', grayscale=True, confidence=0.7))
        
        if lp_error_exist:
            df.at[line, 'Status'] = 'LP não existe!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            bot.sleep(4.5)

            return True
        
    except Exception:
        print('LP exists!')

    try:
        have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
        
        if have_ence:
            df.at[line, 'Status'] = 'Encerrado!'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            press_key('f3', 1)

            bot.sleep(4.5)

            return True
        
    except Exception:
        print('Project don\'t finished yet!')

    try:
        have_lbpa = bot.locateOnScreen('images/LBPA.png', grayscale=True, confidence=0.9)
        
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

# OPEN PROJECT TREE
def open_tree():
    try:
        have_diagram = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
        
        if have_diagram:
            bot.click(46, 232)
            bot.sleep(2)
            bot.click(186, 252)

            bot.sleep(2)

            return False

    except Exception:
        print('Doesn\'t have diagram!')

# CHECK IF THE DIAGRAM IS ALREADY ENCE
def diagram_notLib():
    try:
        have_purchase = list(bot.locateOnScreen('images/ARROW.png', grayscale=True, confidence=0.8, region=arrowCoords))
        
        if have_purchase:
            try:
                have_ence = bot.locateOnScreen('images/ENCE.png', grayscale=True, confidence=0.9)
                
                if have_ence:
                    bot.click(150, 15)
                    bot.click(240, 75)
                    bot.click(520, 270)
                    bot.click(680, 300)
                    bot.sleep(2)

            except Exception:
                print('Diagram is not ENCE!')

            try:
                have_ente = bot.locateOnScreen('images/ENTE.png', grayscale=True, confidence=0.9)
                
                if have_ente:
                    bot.click(150, 15)
                    bot.click(240, 75)
                    bot.click(520, 200)
                    bot.click(680, 240)
                    bot.sleep(2)

            except Exception:
                print('Diagram is not ENTE!')

            return True

    except Exception:
        print('Doesn\'t have any purchase line!')

# FINISH THE TREE LINE STATUS WITH (ENCE)
def finish_treeLine():
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
                        df.at[line, 'Status'] = 'Compromisso pendente!'
                        df.to_excel(excel_path, index=False, engine='openpyxl')

                        press_key('tab', 1)
                        press_key('enter', 1)

                        bot.sleep(5)

                        return 1
                
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

                    try:
                        error_exist = list(bot.locateAllOnScreen('images/ERROR.png', grayscale=True, confidence=0.7))

                        if error_exist:
                            df.at[line, 'Status'] = 'Compromisso pendente!'
                            df.to_excel(excel_path, index=False, engine='openpyxl')

                            press_key('tab', 1)
                            press_key('enter', 1)

                            bot.sleep(5)

                            return 1
                    
                    except Exception:
                        print('Doesn\'t have pending commitment!')

            except Exception as e:
                print(f'Error: {e}')

# CHANGE THE PURCHASE LINE STATUS TO (BAIX CFMN CONF)
def change_purchaseLine_status():
    press_key('tab', 2)
    bot.sleep(0.3)
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
    bot.click(1128, 1100)
    bot.sleep(2)

    try:
        have_advertence = bot.locateOnScreen('images/ADVERTENCE.png', grayscale=True, confidence=0.8)
        
        if have_advertence:
            press_key('tab', 1)
            press_key('enter', 1)

            bot.sleep(2)

    except Exception:
        print('Doesn\'t have advertence!')

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
        print('Doesn\'t have any additional information!')

# FINISH THE PURCHASE LINE STATUS
def ence_purchaseLine():
    workedHours = 0

    bot.click(580, 240)
    bot.sleep(2)

    try:
        have_h = list(bot.locateOnScreen('images/H.png', grayscale=True, confidence=0.8, region=hourCoords))

        if have_h:
            try:
                have_workCenter = list(bot.locateOnScreen('images/FF78012.png', grayscale=True, confidence=0.9, region=workCenterCoords))

                if have_workCenter:
                    print('Doesn\'t have worked hours apointment line!')

            except Exception:
                workedHours = 1

    except Exception:
        try:
            have_workCenter = list(bot.locateOnScreen('images/FF78012.png', grayscale=True, confidence=0.9, region=workCenterCoords))

            if have_workCenter:
                print('Doesn\'t have worked hours apointment line!')

        except Exception:
            workedHours = 1

    bot.click(484, 992)
    bot.sleep(2)
    
    bot.moveTo(606, 956)
    bot.mouseDown()
    bot.moveTo(700, 956, duration=0.25)
    bot.mouseUp()

    bot.sleep(2)

    if workedHours != 1:
        try:
            check_box = list(bot.locateAllOnScreen('images/CHECK.png', grayscale=True, confidence=0.8))
                                    
            if check_box:
                try:
                    have_baixa = list(bot.locateAllOnScreen('images/BAIXCFMN.png', grayscale=True, confidence=0.8))
                    
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
                            bot.click(484, 992)
                            bot.sleep(2)
                            bot.click(600, 992)
                            bot.sleep(2)
                            press_key('enter', 1)
                            bot.sleep(2)
                            press_key('enter', 1)

                            bot.sleep(2)
                            
                            warning_exist = False

                            while not warning_exist:
                                change_purchaseLine_status()

                                try:
                                    warning = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                                    conforder = list(bot.locateAllOnScreen('images/CONFORDER.png', grayscale=True, confidence=0.8))
                            
                                    if warning or conforder:
                                        warning_exist = True

                                except Exception as e:
                                    print(f'Error: {e}')

                            press_key('tab', 2)
                            press_key('enter', 1)
                            bot.sleep(2)
                            press_key('tab', 1)
                            press_key('enter', 1)

                            bot.sleep(2)

                except Exception as e:
                    bot.click(600, 992)
                    bot.sleep(2)

                    warning_exist = False

                    while not warning_exist:
                        change_purchaseLine_status()
                                
                        try:
                            warning = list(bot.locateAllOnScreen('images/WARNING.png', grayscale=True, confidence=0.8))
                            conforder = list(bot.locateAllOnScreen('images/CONFORDER.png', grayscale=True, confidence=0.8))
                    
                            if warning or conforder:
                                warning_exist = True

                        except Exception as e:
                            print(f'Error: {e}')

                    press_key('tab', 2)
                    press_key('enter', 1)
                    bot.sleep(2)
                    press_key('tab', 1)
                    press_key('enter', 1)

                    bot.sleep(2)

        except Exception as e:
            print(f'Error: {e}')
    
    else:
        try:
            check_box = list(bot.locateAllOnScreen('images/CHECK.png', grayscale=True, confidence=0.8))
                                    
            if check_box:
                if len(check_box) != 1:
                    df.at[line, 'Status'] = 'Possui linhas de compra e apontamento!'
                    df.to_excel(excel_path, index=False, engine='openpyxl')

                    bot.sleep(5)
        
                    return 1
                
        except Exception:
            print('Only the apointment line!')
 
    bot.click(580, 200)
    bot.sleep(2)

# CHANGE THE LP STATUS TO ENCE AND SAVE EVERYTHING
def conclusion():
    df.at[line, 'Status'] = 'Encerrado!'
    df.to_excel(excel_path, index=False, engine='openpyxl')

    bot.hotkey('ctrl', 's')

    bot.sleep(5)

# PUT THE ERROR STATUS OF LP AND DON'T SAVE THE CHANGES
def error_conclusion():
    press_key('f3', 1)
    bot.sleep(2)

    try:
        have_changes = list(bot.locateAllOnScreen('images/SAVE.png', grayscale=True, confidence=0.8))

        if have_changes:
            press_key('tab', 1)
            press_key('enter', 1)

    except Exception:
        print('Doesn\'t have changes!')

    bot.sleep(5)

# EXCEL CONFIG
lp_qty = 60
line = 25

# REPEAT QUANTITY TO PROGRAM RUN
repeat_qty = lp_qty - line

# MAIN PROGRAM
for _ in range(repeat_qty):
    jump_all_process = False
    jump_main_function = False
    purchase_line = False
    diagram = True

    pending = 0

    open_project()

    jump_all_process = project_status()

    if not jump_all_process:
        diagram = open_tree()

        if diagram:
            finish_treeLine()
            jump_main_function = True

        if not jump_main_function:
            purchase_line = diagram_notLib()
            
            if purchase_line:
                pending = ence_purchaseLine()

            if pending != 1:
                pending = finish_treeLine()

                if pending != 1:
                    bot.click(186, 212)
                    bot.sleep(2)
                    
                    pending = finish_treeLine()
                    conclusion()

                else:
                    error_conclusion()

            else:
                error_conclusion()
            
    line += 1