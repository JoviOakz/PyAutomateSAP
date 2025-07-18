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

order_qty = 0
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
        heijunka_position = bot.locateOnScreen('images/HEIJUNKA.png', grayscale=True, confidence=0.9)

        if heijunka_position:
            bot.click(bot.center(heijunka_position))
            bot.sleep(1.15)

            press_key('tab', 3)

            for _ in range(repeat_count):
                global line

                order = df.at[line, 'Ordem']

                bot.typewrite(str(order))
                press_key('enter', 1)

                bot.sleep(0.4)

                try:
                    actions_position = bot.locateOnScreen('images/ACTIONS.png', grayscale=True, confidence=0.9)

                    if actions_position:
                        right_x = actions_position.left + (8.75 * (actions_position.width / 10))
                        middle_y = actions_position.top + (actions_position.height / 2)

                        bot.click(right_x, middle_y)
                        bot.sleep(0.65)

                        try:
                            schedule_position = bot.locateOnScreen('images/SCHEDULE.png', grayscale=True, confidence=0.9)

                            if schedule_position:
                                bot.click(bot.center(schedule_position))
                                bot.sleep(0.1)
                                press_key('enter', 1)
                                bot.sleep(0.65)

                                try:
                                    save_position = bot.locateOnScreen('images/SAVE.png', grayscale=True, confidence=0.9)

                                    if save_position:
                                        bot.click(bot.center(save_position))
                                        bot.sleep(0.85)

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
        closure_position = bot.locateOnScreen('images/CLOSURE.png', grayscale=True, confidence=0.9)

        if closure_position:
            bot.click(bot.center(closure_position))
            bot.sleep(1.15)

            enough = False

            while not enough:
                try:
                    finish_position = bot.locateOnScreen('images/FINISH.png', grayscale=True, confidence=0.9)

                    if finish_position:
                        middle_x = finish_position.left + (finish_position.width / 2)
                        threeQ_y = finish_position.top + (3 * (finish_position.height / 4))

                        bot.click(middle_x, threeQ_y)

                        press_key('tab', 1)
                        press_key('enter', 1)
                        bot.sleep(1.85)
                        press_key('enter', 1)

                except Exception:
                    enough += True

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