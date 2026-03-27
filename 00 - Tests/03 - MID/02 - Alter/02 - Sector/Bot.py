# ===== LIBRARIES =====

import pyautogui as bot

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.15

# ===== INITIAL ACTION =====

bot.click(1802, 14)
bot.sleep(1)

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        else:
            bot.press(key)

def find_palletDuBico(press_qty):
    press_key('tab', press_qty)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(0.85)
    press_key('down', 9)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(0.85)
    press_key('tab', 1)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(0.85)
    press_key('down', 34)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(0.85)
    press_key('tab', 3)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(3.5)

def find_orange_color():
    screenshot = bot.screenshot()
    img = screenshot.convert('RGB')
    width, height = img.size
    min_orange = (250, 145, 45)
    max_orange = (255, 155, 55)

    orange_pixels = []

    for x in range(0, width, 5):
        for y in range(0, height, 5):
            r, g, b = img.getpixel((x, y))
            if (min_orange[0] <= r <= max_orange[0] and
                min_orange[1] <= g <= max_orange[1] and
                min_orange[2] <= b <= max_orange[2]):
                orange_pixels.append((x, y))

    if not orange_pixels:
        return None

    avg_x = sum(p[0] for p in orange_pixels) // len(orange_pixels)
    avg_y = sum(p[1] for p in orange_pixels) // len(orange_pixels)

    return (avg_x, avg_y)

def nexist_result():
    press_key('f3', 1)
    bot.sleep(0.5)
    press_key('backspace', 1)
    bot.sleep(0.5)
    bot.typewrite('Nenhum resultado encontrado')
    bot.sleep(0.75)
    
    found = find_orange_color()

    bot.sleep(0.25)
    press_key('esc', 1)
    bot.sleep(0.25)
    
    if found:
        return True

def find_sector_camp():
    press_key('f3', 1)
    bot.sleep(0.5)
    press_key('backspace', 1)
    bot.sleep(0.5)
    bot.typewrite('DU - BICO')
    bot.sleep(0.75)

    orange_center = find_orange_color()

    bot.sleep(0.25)
    press_key('esc', 1)
    bot.sleep(0.25)
    
    if orange_center:
        bot.sleep(0.25)
        bot.click(orange_center)
    else:
        print('Error: DU-BICO don\'t found!')

def change_to_impact():
    press_key('tab', 3)
    bot.sleep(1)
    press_key('space', 1)
    bot.sleep(3.5)

    find_sector_camp()

    press_key('ctrla', 1)
    bot.sleep(0.85)
    press_key('space', 1)
    bot.sleep(0.85)
    press_key('down', 38)
    bot.sleep(0.85)
    press_key('enter', 1)
    bot.sleep(1.25)
    press_key('enter', 1)
    bot.sleep(3.5)
    press_key('enter', 1)
    bot.sleep(1.25)

# ===== MAIN =====

def main():
    press_qty = 18

    while True:
        bot.sleep(1)

        find_palletDuBico(press_qty)
        press_qty = 21

        try:
            if nexist_result():
                break
            else:
                change_to_impact()

        except Exception as e:
            print(f'Error: {e}')

    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()