import pyautogui as bot

bot.FAILSAFE = True
bot.PAUSE = 0.25

# Define coordenadas para regiões e cliques
coordinates = [
    ((474, 430, 33, 26), (488, 440)),
    ((470, 455, 33, 26), (488, 470)),
    ((605, 455, 33, 26), (622, 470))
]

# Muda o status das linhas de compra
def step2_change_status():
    for region, click_position in coordinates:
        try:
            if bot.locateOnScreen('images/CHECK.png', grayscale=True, confidence=0.7, region=region):
                print('encontrado')
            else:
                raise Exception
        except Exception:
            bot.click(click_position)

try:
    step2_change_status()
except Exception as e:
    print(f'Erro: {e}')
