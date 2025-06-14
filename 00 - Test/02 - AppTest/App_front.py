import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QLineEdit, QMessageBox
)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("RS's Register")
        self.setFixedSize(400, 450)

        icon_path = os.path.join(os.path.dirname(__file__), "User_icon_2.ico")
        self.setWindowIcon(QIcon(icon_path))

        font_label = QFont("Arial", 12)
        font_input = QFont("Arial", 11)

        line_edit_style = """
            QLineEdit {
                border: 2px solid #ccc;
                border-radius: 8px;
                padding: 8px;
                background-color: #f9f9f9;
                color: black;
            }
            QLineEdit:focus {
                border: 2px solid #0078d7;
                background-color: #ffffff;
            }
        """

        button_style = """
            QPushButton {
                background-color: #0078d7;
                color: white;
                border-radius: 8px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """

        self.label_lp = QLabel("Insira a LP:")
        self.label_lp.setFont(font_label)
        self.input_lp = QLineEdit()
        self.input_lp.setFont(font_input)
        self.input_lp.setPlaceholderText("Ex.: LP-XXXXXX")
        self.input_lp.setFixedHeight(40)
        self.input_lp.setStyleSheet(line_edit_style)

        self.label_pac = QLabel("Perfil de apropriação de custos:")
        self.label_pac.setFont(font_label)
        self.input_pac = QLineEdit()
        self.input_pac.setFont(font_input)
        self.input_pac.setPlaceholderText("Ex.: ZPS001")
        self.input_pac.setFixedHeight(40)
        self.input_pac.setStyleSheet(line_edit_style)

        self.label_sca = QLabel("Esquema de alocação:")
        self.label_sca.setFont(font_label)
        self.input_sca = QLineEdit()
        self.input_sca.setFont(font_input)
        self.input_sca.setPlaceholderText("Ex.: 05")
        self.input_sca.setFixedHeight(40)
        self.input_sca.setStyleSheet(line_edit_style)

        self.button = QPushButton("Enviar")
        self.button.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        self.button.setFixedHeight(45)
        self.button.setStyleSheet(button_style)
        self.button.clicked.connect(self.send_values)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.addWidget(self.label_lp)
        layout.addWidget(self.input_lp)
        layout.addWidget(self.label_pac)
        layout.addWidget(self.input_pac)
        layout.addWidget(self.label_sca)
        layout.addWidget(self.input_sca)
        layout.addStretch()
        layout.addWidget(self.button)
        self.setLayout(layout)

    def send_values(self):
        lp = self.input_lp.text().strip()
        applicant = self.input_pac.text().strip()
        receptor = self.input_sca.text().strip()

        if not lp or not applicant or not receptor:
            QMessageBox.warning(self, "Erro", "Preencha todos os campos!")
            return
        
        # Apenas mensagem informativa - sem backend
        QMessageBox.information(self, "Enviado", f"Valores enviados:\nLP: {lp}\nPerfil de apropriação: {applicant}\nEsquema de alocação: {receptor}")

        self.input_lp.clear()
        self.input_pac.clear()
        self.input_sca.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())