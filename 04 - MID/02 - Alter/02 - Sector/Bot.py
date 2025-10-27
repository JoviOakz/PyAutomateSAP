# ===== LIBRARIES =====

import pyautogui as bot

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.001

# ===== INITIAL ACTION =====

bot.click(1802, 14)
bot.sleep(1)

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'shtab':
            bot.hotkey('shift', 'tab')
        else:
            bot.press(key)

def find_palletDuBico():
    press_key('tab', 18)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(0.75)
    press_key('down', 9)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(0.75)
    press_key('tab', 1)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(0.75)
    press_key('down', 37)
    bot.sleep(0.75)
    press_key('space', 1)

def change_to_impact():
    print('a')

# ===== MAIN =====

def main():
    find_palletDuBico()
    # change_to_impact()

    # bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()