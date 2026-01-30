# ===== LIBRARIES =====

from pyrfc import Connection
import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.5

# ===== SAP PARAMETERS (PyRFC) =====

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

# ===== INITIAL ACTION =====

# bot.click(1802, 14)

# ===== EXCEL CONFIGURATION =====

excel_path = '../08 - Excels/Cadastro-LPs.xlsx'
df = pd.read_excel(excel_path, engine='openpyxl')

# ===== FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrlf9':
            bot.hotkey('ctrl', 'f9')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'shtab':
            bot.hotkey('shift', 'tab')
        elif key == 'shf1':
            bot.hotkey('shift', 'f1')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def wbs_element_register(data):
    bot.typewrite(str(data['project']))
    bot.sleep(2)
    press_key('enter')
    bot.sleep(2)



    # ==========================================================================================================
    # essa verificação do tamanho pode ser feito uma unica vez e ser criado uma var para o nome 'novo'
    # if (desc < 12c == norm+desc):
    #     bot.typewrite(norm+desc)

    press_key('tab', 3)

    # if (desc < 12c == norm+desc):
    #     bot.typewrite(norm+desc)
    # ==========================================================================================================



    press_key('ctrlf9', 1)
    bot.sleep(2)
    press_key('down', 2)
    press_key('tab', 1)
    press_key('down', 2)
    press_key('ctrla', 1)



    # ==========================================================================================================
    # verificação do Requerente | QMM = 68540007 e outros
    # bot.typewrite('68540007')
    # ==========================================================================================================



    press_key('up', 3)
    press_key('shtab', 1)
    press_key('right', 1)
    press_key('space', 1)
    bot.sleep(2)
    press_key('tab', 1)



    # ==========================================================================================================
    # inserção das datas
    # bot.typewrite(str(data['date']))
    # ==========================================================================================================



    press_key('up', 3)
    press_key('shtab', 1)
    press_key('right', 3)
    press_key('space', 1)
    bot.sleep(2)
    press_key('tab', 4)



    # ==========================================================================================================
    # inserção dos campos gerais e numéricos
    # bot.typewrite('REQUERENTE')
    # press_key('tab', 2)
    # bot.typewrite('QUANTIDADE')
    # press_key('tab', 1)
    # bot.typewrite('PC')
    # press_key('tab', 3)
    # bot.typewrite('VALOR')
    # press_key('tab', 1)
    # bot.typewrite('BRL')
    # ==========================================================================================================



    press_key('tab', 8)
    press_key('space', 1)
    bot.sleep(2)
    press_key('tab', 1)



    # ==========================================================================================================
    # inserção dos parametros
    # bot.typewrite('ZPS001')
    # press_key('tab', 1)
    # bot.typewrite('19')
    # press_key('f3', 1)
    # bot.sleep(2)
    # ==========================================================================================================



    # ==========================================================================================================
    # inserção do objeto de liquidação
    # press_key('tab', 1)
    # bot.typewrite('CENTRO DE CUSTO')
    # ==========================================================================================================
    
    
    
    press_key('f3', 1)
    bot.sleep(2)
    press_key('f3', 1)
    bot.sleep(2)
    press_key('shf1', 1)
    bot.sleep(2)
    press_key('down', 3)
    press_key('ctrla', 1)



    # ==========================================================================================================
    # inserção das responsabilidades e datas
    # bot.typewrite('REQUERENTE')
    # press_key('down', 1)
    # bot.typewrite('DATA')
    # press_key('down', 1)
    # bot.typewrite('DATA')

    # VERIFICAR COMO FAZER PARA CLICAR EM PROCESSAR -> STATUS -> LIBERAR | DE FORMA DINÂMICA!
    # ==========================================================================================================



    press_key('ctrls', 1)
    bot.sleep(3)

def network_creation_via_PyRFC():
    try:
        conn = Connection(**sap_conn_params)
        print('✅ SAP connected successfully')

        # ===== CREATE NETWORK =====
        try:
            # ===== BUILD NETWORK HEADER =====
            network_header = {
                'NETWORK': 'LP053617',
                'PROFILE': 'ZBP0001',
                'NETWORK_TYPE': 'BP01',
                'PLANT': '6854',
                'MRP_CONTROLLER': 'I49',
                'SHORT_TEXT': '4316400030-0229 - BASE - CONSERTO',
                'PROJECT_DEFINITION': 'LP-053617',
                'WBS_ELEMENT': 'LP-053617',
                'START_DATE': '20251119',
                'FINISH_DATE': '20260306'
            }

            # ===== BUILD METHOD TABLE (MANDATORY) =====
            network_method = {
                'REFNUMBER': '000001',
                'OBJECTTYPE': 'NETWORK',
                'METHOD': 'CREATE',
                'OBJECTKEY': 'LP053617'
            }

            conn.call("BAPI_PS_INITIALIZATION")

            # ===== CALL BAPI =====
            resp = conn.call(
                'BAPI_NETWORK_MAINTAIN',
                I_NETWORK=[network_header],
                I_ACTIVITY=[],
                I_METHOD_PROJECT=[network_method]
            )

            print("\n===== RAW SAP RESPONSE =====")
            for key, value in resp.items():
                print(f"\n--- {key} ---")
                print(value)

            # ===== NORMALIZE RETURN =====
            return_raw = resp.get("RETURN", [])

            if isinstance(return_raw, dict):
                return_messages = [return_raw]
            elif isinstance(return_raw, list):
                return_messages = return_raw
            else:
                return_messages = []

            print("\n📩 SAP RETURN MESSAGES:")
                
            for msg in return_messages:
                if isinstance(msg, dict):
                    print(f"[{msg.get('TYPE', '?')}] {msg.get('MESSAGE', '')}")
                else:
                    print(f"[?] {msg}")

            # ===== CHECK ERRORS SAFELY =====
            has_error = any(
                isinstance(msg, dict) and msg.get("TYPE") in ("E", "A")
                for msg in return_messages
            )

            if has_error:
                print("\n❌ ERROR: Network was NOT created due to SAP errors.")

                return {
                    'success': False,
                    'messages': return_messages
                }

            # ========== COMMIT ==========
            conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
            conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
            print("\n✅ SUCCESS: Network created successfully!")

            return {
                'success': True,
                'messages': return_messages
            }

        except Exception as e:
            print(f'\n❌ SAP connection or execution failed\nError: {e}')
            
            return {
                'success': False,
                'messages': [{'TYPE': 'X', 'MESSAGE': str(e)}]
            }

    except Exception as e:
        print(f'Message: SAP connection failed\nError: {e}')

# ===== PROGRAM CONFIGURATION =====

lp_qty = 1
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        # wbs_element_register(data)
        network_creation_via_PyRFC()

        line += 1

    # bot.alert(title='BotText', text='Programa encerrado!')