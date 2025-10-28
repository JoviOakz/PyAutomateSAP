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
    press_key('down', 34)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(0.75)
    press_key('tab', 3)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(1)

def change_to_impact():
    press_key('tab', 3)
    bot.sleep(1)
    press_key('space', 1)
    bot.sleep(2)
    press_key('tab', 25)
    bot.sleep(0.75)
    press_key('space', 1)
    bot.sleep(0.75)
    press_key('down', 38)
    bot.sleep(0.75)
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('enter', 1)
    bot.sleep(1.25)

def find_orange():
    press_key('f3', 1)
    bot.sleep(0.5)
    press_key('backspace', 1)
    bot.sleep(0.5)
    bot.typewrite('Nenhum resultado encontrado')
    bot.sleep(0.75)

    screenshot = bot.screenshot()
    img = screenshot.convert('RGB')
    width, height = img.size
    min_orange = (250, 145, 45)
    max_orange = (255, 155, 55)

    bot.sleep(0.25)
    press_key('esc', 1)

    for x in range(0, width, 5):
        for y in range(0, height, 5):
            r, g, b = img.getpixel((x, y))
            if (min_orange[0] <= r <= max_orange[0] and
                min_orange[1] <= g <= max_orange[1] and
                min_orange[2] <= b <= max_orange[2]):
                return True
            
    return False

# ===== MAIN =====

def main():
    while True:
        bot.sleep(1)

        try:
            if find_orange():
                break
            
        except Exception:
            find_palletDuBico()
            change_to_impact()

    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()