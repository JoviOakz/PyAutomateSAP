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
        self.setFixedSize(400, 400)

        # Definindo o ícone
        icon_path = os.path.join(os.path.dirname(__file__), "User_icon_2.ico")
        self.setWindowIcon(QIcon(icon_path))

        # Label e campo para LP
        self.label_lp = QLabel("Insira a LP:")
        self.label_lp.setFont(QFont("Arial", 12))
        self.label_lp.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.input_lp = QLineEdit()
        self.input_lp.setFont(QFont("Arial", 11))
        self.input_lp.setPlaceholderText("Ex.: LP-XXXXXX")
        self.input_lp.setFixedHeight(35)

        # Label e campo para CCReq
        self.label_ccreq = QLabel("CC do requerente:")
        self.label_ccreq.setFont(QFont("Arial", 12))
        self.label_ccreq.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.input_ccreq = QLineEdit()
        self.input_ccreq.setFont(QFont("Arial", 11))
        self.input_ccreq.setPlaceholderText("Ex.: 685864")
        self.input_ccreq.setFixedHeight(35)

        # Label e campo para Custo
        self.label_cost = QLabel("Insira o custo:")
        self.label_cost.setFont(QFont("Arial", 12))
        self.label_cost.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.input_cost = QLineEdit()
        self.input_cost.setFont(QFont("Arial", 11))
        self.input_cost.setPlaceholderText("Ex.: 4000,00")
        self.input_cost.setFixedHeight(35)

        # Botão
        self.button = QPushButton("Enviar")
        self.button.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        self.button.setFixedHeight(40)
        self.button.clicked.connect(self.send_values)
        self.button.setDefault(True)

        # Conectar "Enter" dos campos
        self.input_lp.returnPressed.connect(self.send_values)
        self.input_ccreq.returnPressed.connect(self.send_values)
        self.input_cost.returnPressed.connect(self.send_values)

        # Layout
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)

        layout.addWidget(self.label_lp)
        layout.addWidget(self.input_lp)
        layout.addWidget(self.label_ccreq)
        layout.addWidget(self.input_ccreq)
        layout.addWidget(self.label_cost)
        layout.addWidget(self.input_cost)
        layout.addWidget(self.button)

        self.setLayout(layout)

    def send_values(self):
        lp = self.input_lp.text()
        ccreq = self.input_ccreq.text()
        cost = self.input_cost.text()

        print(f"LP enviado: {lp}")
        print(f"CC do requerente enviado: {ccreq}")
        print(f"Custo enviado: {cost}")

        # Limpar os campos
        self.input_lp.clear()
        self.input_ccreq.clear()
        self.input_cost.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())