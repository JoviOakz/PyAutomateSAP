import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.35

try:
    ence_location = list(bot.locateOnScreen('../images/ENTE.png', grayscale=True, confidence=0.9))
    if ence_location:
        print('TA ENCERRADO')
except Exception:
    print('ERRO')