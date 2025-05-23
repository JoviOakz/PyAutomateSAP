import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QLineEdit
)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("RS's Register")
        self.setFixedSize(400, 250)

        # Definindo o ícone
        icon_path = os.path.join(os.path.dirname(__file__), "User_icon_2.ico")
        self.setWindowIcon(QIcon(icon_path))

        # Label
        self.label = QLabel("Digite o Part Number:")
        self.label.setFont(QFont("Arial", 12))
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Campo de entrada
        self.input_field = QLineEdit()
        self.input_field.setFont(QFont("Arial", 11))
        self.input_field.setPlaceholderText("Ex.: ABC123")
        self.input_field.setFixedHeight(35)
        self.input_field.returnPressed.connect(self.send_part_number)

        # Botão
        self.button = QPushButton("Enviar")
        self.button.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        self.button.setFixedHeight(40)
        self.button.clicked.connect(self.send_part_number)
        self.button.setDefault(True)

        # Layout
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)

        layout.addWidget(self.label)
        layout.addWidget(self.input_field)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def send_part_number(self):
        part_number = self.input_field.text()
        print(f"Part Number enviado: {part_number}")
        self.input_field.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    # Definindo também o ícone no QApplication
    icon_path = os.path.join(os.path.dirname(__file__), "icone.ico")
    app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())