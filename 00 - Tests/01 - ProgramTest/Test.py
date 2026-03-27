# ==============================================================================================================

# import pyautogui

# pyautogui.alert(text='Fala gurias eu sou o Ribeiro!', title='Caixa de texto!', button='OK')
# pyautogui.confirm(text='Fala gurias eu sou o Ribeiro!', title='Caixa de texto!', buttons=['SIM', 'OR NOT?!'])
# print(pyautogui.prompt(text='Diga a sua cor favorita', title='Caixa de pergunta', default='Hobs'))
# print(pyautogui.password(text='Diga a sua senha favorita', title='Caixa de phishing', default='123', mask='*'))

#! python3
# import pyautogui, sys
# print('Press Ctrl-C to quit.')
# try:
#     while True:
#         x, y = pyautogui.position()
#         positionStr = 'X: ' + str(x).rjust(4) + ' Y: ' + str(y).rjust(4)
#         print(positionStr, end='')
#         print('\b' * len(positionStr), end='', flush=True)
# except KeyboardInterrupt:
#     print('')

# pyautogui.moveTo(100, 100, 2, pyautogui.easeInBounce)
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInElastic)
# pyautogui.moveTo(100, 100, 2, pyautogui.easeOutQuad)
# pyautogui.moveTo(100, 100, 2, pyautogui.easeInQuad)


# pyautogui.sleep(3)

# for _ in range(100):
#     pyautogui.click()

# pyautogui.click(clicks=10000, interval=0.00001)
# pyautogui.click(clicks=10000)

# ==============================================================================================================

# import pyautogui as bot
# import pandas as pd
# import pyperclip

# bot.FAILSAFE = True
# bot.PAUSE = 1

# bot.click(1162, 1172)

# df = pd.read_excel("Pasta1.xlsx", engine='openpyxl')

# group = []
# line = 0

# for _ in range(15):
#     value = df.at[line, "Coluna"]
#     group.append(value)
#     line += 1

# all_values = '\n'.join(map(str, group))

# pyperclip.copy(all_values)
# bot.hotkey('ctrl', 'v')
    
# ==============================================================================================================

import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 1

bot.click(1802, 14)

bot.sleep(1)

try:
    save_position = bot.locateOnScreen('SAVE.png', grayscale=True, confidence=0.9)

    if save_position:
        bot.moveTo(bot.center(save_position))

except Exception as e:
    print(f'Error: Save not found!\nException: {e}')