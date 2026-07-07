# 📦 Instalação de Bibliotecas Python - Projetos de Automatização de Processos e Interface

Este projeto utiliza diversas bibliotecas Python voltadas para automação de tarefas, manipulação de arquivos, visão computacional e criação de interfaces gráficas.

Abaixo estão listadas as bibliotecas necessárias com instruções de instalação.

---

## 🐍 Instalação do Python

- **Python 3.12 (64 bits)** — recomendado para maior compatibilidade com PyRFC  
  🔗 [Download Python 3.12.10](https://www.python.org/downloads/release/python-31210/)

- Atualizar pip  
  ```bash
  python -m pip install --upgrade pip
  ```

---

## 🔧 Dependências Externas

Estas ferramentas devem ser instaladas separadamente, pois não estão disponíveis diretamente via `pip`.

- Baixar [Poppler v24.08.0-0](https://github.com/oschwartz10612/poppler-windows/releases/tag/v24.08.0-0)  
- Repositório [Tesseract OCR (oficial)](https://github.com/tesseract-ocr/tesseract)  
- Releases [PyRFC (SAP) - Versão 3.3.1](https://github.com/SAP-archive/PyRFC/releases) -> compatível com python 3.12

Como complemento do PyRFC:<br>
Seguir o caminho | 02 - Others -> 99 - Files | para adquirir o arquivo SDK do SAP -> nwrfcsdk.zip

---

## 📚 Bibliotecas Python

Instale todas as bibliotecas abaixo com o seguinte comando:

```bash
pip install opencv-python openpyxl pandas pyautogui PyQt6 pdf2image pytesseract cython pyperclip oracledb
```

---

## 🚀 Próximos Passos

🔹 **Atualizar bot de encerramento de LPs**
  - Encerramento completo e efetivo de quaisquer tipos de status em LPs

🔹 **Atualizar bot de apontamento de horas**
  - Realizar o apontamento automaticamente conforme regras de negócio