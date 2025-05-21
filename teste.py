import pyautogui

# pyautogui.alert('Dá um OK aí!')
# print(pyautogui.prompt('Escreva algo'))

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


pyautogui.sleep(3)

# for _ in range(100):
#     pyautogui.click()

pyautogui.click(clicks=10000, interval=0.00001)
# pyautogui.click(clicks=10000)