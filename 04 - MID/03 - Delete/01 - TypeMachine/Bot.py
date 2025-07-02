# ===== LIBRARIES =====

import pyautogui as bot

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.1

# ===== INITIAL ACTION =====

bot.click(1802, 14)
bot.sleep(0.5)

# ===== PROGRAM CONFIGURATION =====

part_number_qty = 10
line = 0
repeat_count = part_number_qty - line

# ===== FUNCTIONS =====

def process_part_numbers():
    for _ in range(repeat_count):
        bot.click(710, 1046)
        bot.press('tab')
        bot.press('enter')

# ===== MAIN =====

def main():
    process_part_numbers()
    bot.alert(title='BotText', text='Programa encerrado!')

if __name__ == '__main__':
    main()