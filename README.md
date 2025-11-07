# 📦 Instalação de Bibliotecas Python - Projeto de Automação e Interface

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

Como complemento do PyRFC, seguir o caminho | 99 - Others -> 99 - Files | para adquirir o arquivo SDK do SAP -> nwrfcsdk.zip

---

## 📚 Bibliotecas Python

Instale todas as bibliotecas abaixo com o seguinte comando:

```bash
pip install opencv-python openpyxl pandas pyautogui PyQt6 pdf2image pytesseract cython
```

---

## 🚀 Próximos Passos

🔹 **Atualizar bot para apontamento de horas**  
  - Verificar linhas em que é possível criar o apontamento  
  - Realizar o apontamento automaticamente conforme regras de negócio  

🔹 **Criar novo bot para cadastros de LP**  
  - Utilizar **PyRFC** para resgate de informações diretamente do SAP  
  - Implementar lógica de validação e tratamento de dados antes do envio