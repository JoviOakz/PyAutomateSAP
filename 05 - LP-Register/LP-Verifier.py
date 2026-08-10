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
        if key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrll':
            bot.hotkey('ctrl', 'l')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlup':
            bot.hotkey('ctrl', 'up')
        else:
            bot.press(key)

def verifier(repeat_qty):
    line = 0

    for _ in range(repeat_qty):
        line += 1
        press_key('ctrlc', 1)
        press_key('ctrll', 1)
        press_key('ctrlv', 1)
        press_key('enter', 1)
        press_key('esc', 1)
        press_key('tab', 1)
        bot.typewrite('CADASTRAR')
        press_key('tab', 1)

        if line != repeat_qty:
            press_key('ctrlup', 2)
        else:
            press_key('ctrlup', 1)
            
        press_key('del', 1)
        press_key('down', 1)

# ===== PROGRAM CONFIGURATION =====

repeat_qty = 32

# ===== MAIN =====

if __name__ == '__main__':
    verifier(repeat_qty)

    bot.alert(title='BotText', text='Program successfully completed')