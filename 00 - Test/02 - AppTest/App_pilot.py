import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QLabel, QPushButton, QVBoxLayout, QWidget, QLineEdit, QMessageBox
)
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtCore import Qt
from pyrfc import Connection

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

        self.label_ccreq = QLabel("Objeto de liquidação:")
        self.label_ccreq.setFont(font_label)
        self.input_ccreq = QLineEdit()
        self.input_ccreq.setFont(font_input)
        self.input_ccreq.setPlaceholderText("Ex.: 685864")
        self.input_ccreq.setFixedHeight(40)
        self.input_ccreq.setStyleSheet(line_edit_style)

        self.label_cost = QLabel("Insira o Custo:")
        self.label_cost.setFont(font_label)
        self.input_cost = QLineEdit()
        self.input_cost.setFont(font_input)
        self.input_cost.setPlaceholderText("Ex.: 4000,00")
        self.input_cost.setFixedHeight(40)
        self.input_cost.setStyleSheet(line_edit_style)

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
        layout.addWidget(self.label_ccreq)
        layout.addWidget(self.input_ccreq)
        layout.addWidget(self.label_cost)
        layout.addWidget(self.input_cost)
        layout.addStretch()
        layout.addWidget(self.button)
        self.setLayout(layout)

    def validate_inputs(self):
        """Valida os campos de entrada antes de enviar ao SAP"""
        lp = self.input_lp.text().strip()
        applicant = self.input_ccreq.text().strip()
        receptor = self.input_cost.text().strip()

        if not lp or not applicant or not receptor:
            return False, "Preencha todos os campos!"

        if not lp.startswith('LP-') or len(lp) != 9 or not lp[3:].isdigit():
            return False, "Formato de LP inválido. Use LP-XXXXXX (6 dígitos)"

        if not applicant.isdigit():
            return False, "Objeto de liquidação deve conter apenas números"

        try:
            # Converte o formato brasileiro para float (4000,00 -> 4000.00)
            custo = float(receptor.replace('.', '').replace(',', '.'))
            if custo <= 0:
                return False, "O custo deve ser maior que zero"
        except ValueError:
            return False, "Custo inválido. Use formato 4000,00"

        return True, ""

    def send_values(self):
        """Envia os valores para o SAP após validação"""
        validation, message = self.validate_inputs()
        if not validation:
            QMessageBox.warning(self, "Erro", message)
            return

        lp = self.input_lp.text().strip()
        applicant = self.input_ccreq.text().strip()
        receptor = self.input_cost.text().replace('.', '').replace(',', '.')

        try:
            sap_connection_params = {
                'user': 'MAO8CT',
                'passwd': '86IQ3J$.7vCj@',
                'ashost': 'rb3ps0a0.server.bosch.com',
                'sysnr': '00',
                'client': '011',
                'lang': 'PT',
                'timeout': 30  # Adicionado timeout
            }

            with Connection(**sap_connection_params) as conn:  # Usando context manager
                # Verifica se a LP existe no SAP
                lp_check = conn.call('RFC_READ_TABLE',
                                   QUERY_TABLE='PRPS',
                                   FIELDS=[{'FIELDNAME': 'PSPNR'}],
                                   OPTIONS=[f"PSPNR = '{lp}'"])
                
                if not lp_check['DATA']:
                    raise ValueError(f"LP {lp} não encontrada no SAP")

                # Atualiza o projeto
                conn.call('BAPI_PROJECT_MAINTAIN',
                    I_PROJECT_DEFINITION={
                        'PROJECT_DEFINITION': lp,
                        'DESCRIPTION': 'Alteracao via PyQt6'
                    },
                    I_PROJECT_DEFINITION_UPD={
                        'PROJECT_DEFINITION': 'X',
                        'DESCRIPTION': 'X'
                    },
                    I_METHOD_PROJECT={
                        'METHOD': 'UPDATE'
                    },
                    I_WBS_ELEMENT_TABLE_UPDATE=[{
                        'WBS_ELEMENT': lp,
                        'APPLICANT_NO': applicant
                    }],
                    I_WBS_ELEMENT_TABLE_UPDATE_UPD=[{
                        'WBS_ELEMENT': ' ',
                        'APPLICANT_NO': 'X'
                    }]
                )

                # Obtém o objeto do projeto
                objnr_result = conn.call('RFC_READ_TABLE',
                                         QUERY_TABLE='PRPS',
                                         DELIMITER='|',
                                         FIELDS=[{'FIELDNAME': 'OBJNR'}],
                                         OPTIONS=[f"PSPNR = '{lp}'"])

                objnr = objnr_result['DATA'][0]['WA'].split('|')[0].strip()

                # Lê e atualiza as regras de liquidação
                regra_atual = conn.call('K_SETTLEMENT_RULE_READ', OBJNR=objnr, OBJART='PROJ')
                settlement_rules = regra_atual.get('SETTL_RULE', [])
                
                if not settlement_rules:
                    raise ValueError("Nenhuma regra de liquidação encontrada para este projeto")

                for regra in settlement_rules:
                    regra['EMPGE'] = receptor

                conn.call('K_SETTLEMENT_RULE_CHECK', OBJNR=objnr, SETTL_RULE=settlement_rules)
                conn.call('K_SETTLEMENT_RULE_SAVE', OBJNR=objnr, SETTL_RULE=settlement_rules)
                conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')

                QMessageBox.information(self, "Sucesso", "Dados enviados com sucesso ao SAP!")

        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao conectar ou enviar ao SAP:\n{str(e)}")
            # Aqui você poderia adicionar um log do erro
        finally:
            self.input_lp.clear()
            self.input_ccreq.clear()
            self.input_cost.clear()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())