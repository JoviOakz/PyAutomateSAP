# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.15

# ===== INITIAL ACTION =====

bot.click(1802, 14)
bot.sleep(0.5)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== PROGRAM CONFIGURATION =====

order_qty = 2
line = 0
repeat_count = order_qty - line

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def demand_conclusion():
    try:
        heijunka_position = bot.screenshot(region=(52, 364, 122, 90))
        heijunka_founded = bot.locateOnScreen(heijunka_position, grayscale=True, confidence=0.9)

        if heijunka_founded:
            bot.click(bot.center(heijunka_founded))
            bot.sleep(1.15)

            press_key('tab', 3)

            for _ in range(repeat_count):
                global line

                order = df.at[line, 'Ordem']

                bot.typewrite(str(order))
                press_key('enter', 1)

                bot.sleep(0.4)

                try:
                    actions_position = bot.screenshot(region=(1642, 484, 170, 38))
                    actions_founded = bot.locateOnScreen(actions_position, grayscale=True, confidence=0.9)

                    if actions_founded:
                        right_x = actions_founded.left + (8.75 * (actions_founded.width / 10))
                        middle_y = actions_founded.top + (actions_founded.height / 2)

                        bot.click(right_x, middle_y)
                        bot.sleep(0.65)

                        try:
                            schedule_position = bot.screenshot(region=(744, 606, 44, 38))
                            schedule_founded = bot.locateOnScreen(schedule_position, grayscale=True, confidence=0.9)

                            if schedule_founded:
                                bot.click(bot.center(schedule_founded))
                                bot.sleep(0.1)
                                press_key('enter', 1)
                                bot.sleep(0.65)

                                try:
                                    save_position = bot.screenshot(region=(1102, 690, 84, 40))
                                    save_founded = bot.locateOnScreen(save_position, grayscale=True, confidence=0.9)

                                    if save_founded:
                                        bot.click(bot.center(save_founded))
                                        bot.sleep(0.9)

                                        press_key('enter', 1)
                                        bot.sleep(0.4)

                                except Exception as e:
                                    print(f'Error: Save not found!\nException: {e}')
                                
                        except Exception as e:
                            print(f'Error: Schedule not found!\nException: {e}')

                        bot.click(right_x, middle_y - 50)
                        press_key('shtab', 7)

                except Exception as e:
                    print(f'Error: Action not found!\nException: {e}')

                line += 1
            
    except Exception as e:
        print(f'Error: Heijunka not found!\nException: {e}')

def close_order():
    try:
        closure_position = bot.screenshot(region=(22, 570, 186, 116))
        closure_founded = bot.locateOnScreen(closure_position, grayscale=True, confidence=0.9)

        if closure_founded:
            bot.click(bot.center(closure_founded))
            bot.sleep(1.15)

            enough = False

            while not enough:
                try:
                    finish_position = bot.screenshot(region=(1696, 426, 70, 102))
                    finish_founded = bot.locateOnScreen(finish_position, grayscale=True, confidence=0.9)

                    if finish_founded:
                        middle_x = finish_founded.left + (finish_founded.width / 2)
                        threeQ_y = finish_founded.top + (3 * (finish_founded.height / 4))

                        bot.click(middle_x, threeQ_y)

                        press_key('tab', 1)
                        press_key('enter', 1)
                        bot.sleep(1.85)
                        press_key('enter', 1)

                except Exception:
                    enough = True

    except Exception as e:
        print(f'Error : Closure not found!\nException: {e}')

# ===== MAIN =====

def main():
    if repeat_count != 0:   
        demand_conclusion()
    
    close_order()
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()