import pyautogui as bot

bot.sleep(2)

# try:
#     not_exist_lp = list(bot.locateAllOnScreen('../images/LPNOTEXIST.png', grayscale=True, confidence=0.9))

#     if not_exist_lp:
#         print('Achou!')

# except Exception as e:
#     print(f'Error: {e}')

try:
    not_exist_lp = list(bot.locateAllOnScreen('../images/NFILLEDLINE.png', grayscale=True, confidence=0.9))

    if not_exist_lp:
        print('Achou!')

except Exception as e:
    print(f'Error: {e}')