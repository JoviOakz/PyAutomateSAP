import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

bot.moveTo(606, 848)
bot.mouseDown()
bot.moveTo(846, 848, duration=0.5)
bot.mouseUp()
bot.sleep(0.3)

try:
    check_box = list(bot.locateAllOnScreen('../images/CHECK.png', grayscale=True, confidence=0.8))
    if check_box:
        try:
            have_baixa = list(bot.locateAllOnScreen('../images/BAIXACONF.png', grayscale=True, confidence=0.8))
            if have_baixa:
                if len(check_box) == len(have_baixa):
                    print('Deu boa')
                else:
                    print('Deu merda')
        except Exception as e:
            print(f'Erro: {e}')
except Exception as e:
    print(f'Erro: {e}')