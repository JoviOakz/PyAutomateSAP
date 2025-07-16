# ===== LIBRARIES =====

import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.1

# ===== INITIAL ACTION =====

bot.click(1802, 14)
bot.sleep(0.5)

# ===== EXCEL CONFIGURATION =====

EXCEL_PATH = "Data.xlsx"
df = pd.read_excel(EXCEL_PATH, engine='openpyxl')

# ===== PROGRAM CONFIGURATION =====

order_qty = 2 #38
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
            center = bot.center(heijunka_position)
            bot.click(center)

            bot.sleep(1)

            press_key('tab', 3)

            for _ in range(repeat_count):
                global line

                order = df.at[line, 'Ordem']

                bot.typewrite(str(order))
                press_key('enter', 1)

                bot.sleep(0.3)

                try:
                    actions_position = bot.locateOnScreen('images/ACTIONS.png', grayscale=True, confidence=0.9)

                    if actions_position:
                        right_x = actions_position.left + (8.75 * (actions_position.width / 10))
                        middle_y = actions_position.top + (actions_position.height / 2)

                        bot.moveTo(right_x, middle_y)
                        bot.sleep(0.3)

                        # =============================================================================================

                        # Parte onde clica no OK duplo e confirma a data atual, dá um OK e segue.

                        # =============================================================================================

                        bot.click(right_x, middle_y - 50)
                        press_key('shtab', 7)

                except Exception as e:
                    print(f'Error: {e}')

                line += 1
            
    except Exception as e:
        print(f'Error: {e}')

def close_order():
    print('Hello World!')

# ===== MAIN =====

def main():
    demand_conclusion()
    close_order()
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()