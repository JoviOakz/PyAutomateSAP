import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.click(1802, 14)

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

serie = 6

for _ in range(64):
    press_key('ctrla', 1)
    press_key('right', 1)
    
    if serie < 10:
        press_key('backspace', 2)
    else:
        press_key('backspace', 3)
        
    bot.typewrite(str(serie))
    press_key('enter', 1)

    bot.sleep(1.5)

    if serie < 10:
        var = '0' + str(serie)
    else:
        var = str(serie)

    bot.typewrite('Encosto Agulha CRI, P e Zexel - 0' + var)

    bot.sleep(0.3)

    press_key('ctrls', 1)

    serie += 1

    bot.sleep(3)