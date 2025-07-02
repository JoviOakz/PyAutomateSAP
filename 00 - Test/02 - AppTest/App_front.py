import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QWidget,
    QLineEdit, QMessageBox, QCheckBox
)
from PyQt6.QtGui import QFont, QIcon, QIntValidator
from PyQt6.QtCore import Qt

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("RS's Register")
        self.setFixedSize(400, 650)

        icon_path = os.path.join(os.path.dirname(__file__), "User_icon_2.ico")
        self.setWindowIcon(QIcon(icon_path))

        font_label = QFont("Arial", 12)
        font_input = QFont("Arial", 11)

        line_edit_style = """
            QLineEdit {
                border: 2px solid #ccc;
                border-radius: 6px;
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
                border-radius: 6px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
        """

        self.label_lp = QLabel("LP a ser cadastrada:")
        self.label_lp.setFont(font_label)
        self.input_lp = QLineEdit()
        self.input_lp.setFont(font_input)
        self.input_lp.setPlaceholderText("LP-XXXXXX")
        self.input_lp.setFixedHeight(40)
        self.input_lp.setMaxLength(9)
        self.input_lp.setValidator(QIntValidator(0, 999999))
        self.input_lp.setStyleSheet(line_edit_style)
        self.input_lp.textChanged.connect(self.update_lp_text)

        self.label_debitando = QLabel("Essa RS está debitando em:")
        self.label_debitando.setFont(font_label)

        self.checkbox_cc = QCheckBox("Centro de Custo")
        self.checkbox_pep = QCheckBox("Elemento PEP")
        self.checkbox_bm = QCheckBox("BM")
        self.checkbox_ordem = QCheckBox("Ordem")

        for cb in (self.checkbox_cc, self.checkbox_pep, self.checkbox_bm, self.checkbox_ordem):
            cb.setFont(font_input)
            cb.toggled.connect(self.on_checkbox_toggled)

        self.label_allcsch = QLabel("Esquema de alocação:")
        self.label_allcsch.setFont(font_label)
        self.input_allcsch = QLineEdit()
        self.input_allcsch.setFont(font_input)
        self.input_allcsch.setPlaceholderText("Ex.: 05")
        self.input_allcsch.setFixedHeight(40)
        self.input_allcsch.setMaxLength(2)
        self.input_allcsch.setValidator(QIntValidator(0, 99))
        self.input_allcsch.setStyleSheet(line_edit_style)

        self.label_dbt = QLabel("Débito:")
        self.label_dbt.setFont(font_label)
        self.input_dbt = QLineEdit()
        self.input_dbt.setFont(font_input)
        self.input_dbt.setFixedHeight(40)
        self.input_dbt.setMaxLength(24)
        self.input_dbt.setStyleSheet(line_edit_style)
        self.input_dbt.textChanged.connect(self.update_debit_format)

        self.button = QPushButton("Enviar")
        self.button.setFont(QFont("Arial", 11, QFont.Weight.Bold))
        self.button.setFixedHeight(45)
        self.button.setStyleSheet(button_style)
        self.button.clicked.connect(self.send_values)

        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.addWidget(self.label_lp)
        layout.addWidget(self.input_lp)
        layout.addWidget(self.label_debitando)
        layout.addWidget(self.checkbox_cc)
        layout.addWidget(self.checkbox_pep)
        layout.addWidget(self.checkbox_bm)
        layout.addWidget(self.checkbox_ordem)
        layout.addWidget(self.label_allcsch)
        layout.addWidget(self.input_allcsch)
        layout.addWidget(self.label_dbt)
        layout.addWidget(self.input_dbt)
        layout.addStretch()
        layout.addWidget(self.button)
        self.setLayout(layout)

    def update_lp_text(self, text):
        digits = ''.join(filter(str.isdigit, text))[:6]
        if len(digits) == 6:
            formatted = f"LP-{digits}"
        else:
            formatted = digits
        self.input_lp.blockSignals(True)
        self.input_lp.setText(formatted)
        self.input_lp.blockSignals(False)

    def on_checkbox_toggled(self, checked):
        sender = self.sender()
        if checked:
            for cb in (self.checkbox_cc, self.checkbox_pep, self.checkbox_bm, self.checkbox_ordem):
                if cb != sender:
                    cb.blockSignals(True)
                    cb.setChecked(False)
                    cb.blockSignals(False)

            if sender == self.checkbox_cc:
                self.input_allcsch.setReadOnly(False)
                self.input_allcsch.clear()
                self.input_dbt.setPlaceholderText("Ex.: 685485")
                self.input_dbt.setMaxLength(6)
            elif sender == self.checkbox_pep:
                self.input_allcsch.setText("07")
                self.input_allcsch.setReadOnly(True)
                self.input_dbt.setPlaceholderText("Ex.: LP-011144")
                self.input_dbt.setMaxLength(9)
            elif sender == self.checkbox_bm:
                self.input_allcsch.setText("07")
                self.input_allcsch.setReadOnly(True)
                self.input_dbt.setPlaceholderText("Ex.: BM-00045998_001_00000008")
                self.input_dbt.setMaxLength(24)
            elif sender == self.checkbox_ordem:
                self.input_allcsch.setText("07")
                self.input_allcsch.setReadOnly(True)
                self.input_dbt.setPlaceholderText("Ex.: 68500066947")
                self.input_dbt.setMaxLength(11)
            self.input_dbt.clear()

    def update_debit_format(self, text):
        if self.checkbox_pep.isChecked():
            digits = ''.join(filter(str.isdigit, text))[:6]
            if len(digits) == 6:
                formatted = f"LP-{digits}"
            else:
                formatted = digits
            self.input_dbt.blockSignals(True)
            self.input_dbt.setText(formatted)
            self.input_dbt.blockSignals(False)
        elif self.checkbox_bm.isChecked():
            digits = ''.join(filter(str.isdigit, text))[:19]
            if len(digits) == 19:
                formatted = f"BM-{digits[:8]}_{digits[8:11]}_{digits[11:19]}"
            else:
                formatted = digits
            self.input_dbt.blockSignals(True)
            self.input_dbt.setText(formatted)
            self.input_dbt.blockSignals(False)

    def send_values(self):
        lp = self.input_lp.text().strip()
        debit = self.input_dbt.text().strip()
        allocation_scheme = self.input_allcsch.text().strip()

        if not lp or not debit or not allocation_scheme:
            QMessageBox.warning(self, "Erro", "Preencha todos os campos!")
            return

        if not lp.startswith("LP-") or len(lp) != 9 or not lp[3:].isdigit():
            QMessageBox.warning(self, "Erro", "LP a ser cadastrada deve estar no formato correto: LP-XXXXXX")
            return

        if not (self.checkbox_cc.isChecked() or self.checkbox_pep.isChecked() or
                self.checkbox_bm.isChecked() or self.checkbox_ordem.isChecked()):
            QMessageBox.warning(self, "Erro", "Selecione uma das opções de débito!")
            return

        if self.checkbox_cc.isChecked():
            if len(debit) != 6 or not debit.isdigit():
                QMessageBox.warning(self, "Erro", "Centro de Custo deve ser válido.")
                return

        elif self.checkbox_pep.isChecked():
            if not debit.startswith("LP-") or len(debit) != 9:
                QMessageBox.warning(self, "Erro", "Elemento PEP deve estar no formato correto: LP-XXXXXX.")
                return

        elif self.checkbox_bm.isChecked():
            if not debit.startswith("BM-") or len(debit) != 24 or '_' not in debit:
                QMessageBox.warning(self, "Erro", "BM deve ter o formato correto: BM-XXXXXXXX_XXX_XXXXXXXX")
                return

        elif self.checkbox_ordem.isChecked():
            if len(debit) != 11 or not debit.isdigit():
                QMessageBox.warning(self, "Erro", "Ordem deve ser válida.")
                return

        tipo = "Centro de Custo" if self.checkbox_cc.isChecked() else \
            "Elemento PEP" if self.checkbox_pep.isChecked() else \
            "BM" if self.checkbox_bm.isChecked() else "Ordem"

        QMessageBox.information(self, "Enviado",
                                f"{lp}\nEsquema: {allocation_scheme}\nTipo: {tipo}\nDébito: {debit}")

        self.input_lp.clear()
        self.input_allcsch.clear()
        self.input_allcsch.setReadOnly(False)
        self.input_dbt.clear()
        self.input_dbt.setPlaceholderText("")
        for cb in (self.checkbox_cc, self.checkbox_pep, self.checkbox_bm, self.checkbox_ordem):
            cb.setChecked(False)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())