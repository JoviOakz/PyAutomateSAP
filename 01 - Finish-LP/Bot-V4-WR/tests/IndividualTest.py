# import pyautogui as bot

# bot.FAILSAFE = True
# bot.PAUSE = 0.25

# hourCoords = (880, 332, 50, 188)
# workCenterCoords = (986, 332, 108, 188)

# bot.sleep(1)

# try:
#     have_h = list(bot.locateOnScreen('../images/H.png', grayscale=True, confidence=0.8, region=hourCoords))

#     if have_h:
#         try:
#             have_workCenter = list(bot.locateOnScreen('../images/FF78012.png', grayscale=True, confidence=0.9, region=workCenterCoords))

#             if have_workCenter:
#                 print('Não tem WH')
#         except Exception:
#             print('Tem apontamento')
# except Exception:
#     print('Não tem H')

check_box = 'dois'
have_baixa = 'um'

if len(check_box) == len(have_baixa)+2:
    print('São do mesmo tamanho')
else:
    print(len(check_box))
    print(len(have_baixa))