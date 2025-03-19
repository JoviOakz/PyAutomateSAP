import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

workCenterCoords = (986, 332, 108, 188)

bot.sleep(1)
try:
    have_workCenter = list(bot.locateOnScreen('../images/FF78012.png', grayscale=True, confidence=0.8, region=workCenterCoords))

    # REVISAR QUAL É O VALOR RETORNADO QUANDO NÃO SE ENCONTRA A IMAGEM NA TELA
    if not have_workCenter:
        print(have_workCenter)
    else:
        print(have_workCenter)

except Exception as e:
    print(f'Error: {e}')