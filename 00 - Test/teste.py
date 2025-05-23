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

import sys
from PyQt6.QtWidgets import QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QLineEdit

# Subclasse para nossa janela principal
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("RS's Register")

        self.resize(400, 200)

        # Widgets
        self.label = QLabel("Digite o Part Number:")
        self.input_field = QLineEdit()
        self.button = QPushButton("Enviar")

        # Conectar botão ao método
        self.button.clicked.connect(self.send_part_number)

        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.input_field)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def send_part_number(self):
        # Pega o texto digitado
        part_number = self.input_field.text()

        # Imprime no terminal
        print(f"Part Number enviado: {part_number}")

        # (Opcional) Limpar o campo depois de enviar
        self.input_field.clear()

# Execução do aplicativo
app = QApplication(sys.argv)

window = MainWindow()
window.show()

sys.exit(app.exec())

# ==============================================================================================================