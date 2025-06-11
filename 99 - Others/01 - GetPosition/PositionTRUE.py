import pyautogui as bot

try:
    while True:
        print(bot.position())

except KeyboardInterrupt:
    print('')