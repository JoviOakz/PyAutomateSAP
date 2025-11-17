# ===== LIBRARIES =====

from pyrfc import Connection
import pyautogui as bot
import pandas as pd

# ===== GLOBAL SETTINGS =====

bot.FAILSAFE = True
bot.PAUSE = 0.5

# ===== INITIAL ACTION =====

bot.click(1802, 14)

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

def getInformation(pm_value):
    sap_conn_params = {
        'user': 'MAO8CT',
        'passwd': '86IQ3J$.7vCj@',
        'ashost': 'rb3ps0a0.server.bosch.com',
        'sysnr': '00',
        'client': '011',
        'lang': 'PT'
    }

    payee = resp = liquidation_obj = quantity = description = part_number = deliverTo = cost = date = project = ''

    try:
        conn = Connection(**sap_conn_params)

        try:
            result_0054 = conn.call(
                'RFC_READ_TABLE',
                QUERY_TABLE='Z22I0054_MD',
                ROWCOUNT=5,
                OPTIONS=[{
                    'TEXT': f"QMNUM = '{pm_value}'"
                }],
                DELIMITER='|'
            )
        
            if result_0054['DATA']:
                raw_line = result_0054['DATA'][0]['WA']
                fields = raw_line.split('|')

                payee = fields[7].strip()
                resp = fields[9].strip()
                liquidation_obj = fields[12].strip()

        except Exception as e:
            print(f'Message: error with Z22I0054_MD table\nError: {e}')

        try:
            result_0055 = conn.call(
                'RFC_READ_TABLE',
                QUERY_TABLE='Z22I0055_MD',
                ROWCOUNT=5,
                OPTIONS=[{
                    'TEXT': f"QMNUM = '{pm_value}'"
                }],
                DELIMITER='|'
            )
        
            if result_0055['DATA']:
                raw_line = result_0055['DATA'][0]['WA']
                fields = raw_line.split('|')

                quantity = int(float(fields[4].strip()))
                description = fields[5].strip()
                part_number = fields[6].strip()
                deliverTo = fields[9].strip()
                cost = fields[10].strip()
                date = fields[12].strip()
                project = fields[13].strip()

                date = str(date[6:8]) + '.' + str(date[4:6]) + '.' + str(date[:4])

        except Exception as e:
            print(f'Message: error with Z22I0055_MD table\nError: {e}')
        
    except Exception as e:
        print(f'Message: SAP connection attempt failed\nError: {e}')

    return {
        'payee': payee,
        'resp': resp,
        'liquidation_obj': liquidation_obj,
        'quantity': quantity,
        'description': description,
        'part_number': part_number,
        'deliverTo': deliverTo,
        'cost': cost,
        'date': date,
        'project': project
    }

def wbs_element_register(data):
    bot.typewrite(str(data['project']))
    bot.sleep(2)
    press_key('enter')
    bot.sleep(2)



    # ==========================================================================================================
    # essa verificação do tamanho pode ser feito uma unica vez e ser criado uma var para o nome "novo"
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

def network_creation(data):
    press_key('enter', 1)
    


    # ==========================================================================================================
    # INSERIR O NÚMERO DA LP SEM O TRAÇO
    # ==========================================================================================================



    press_key('enter', 1)
    bot.sleep(2)

# ===== PROGRAM CONFIGURATION =====

lp_qty = 1
line = 0
repeat_qty = lp_qty - line

# ===== MAIN =====

if __name__ == '__main__':
    for _ in range(repeat_qty):
        pm_value = df.at[line, 'PM']
        pm_value = str(pm_value).zfill(12)
        data = getInformation(pm_value)

        print(data)

        wbs_element_register(data)

        # ===================================================================================================
        # VERIFICAR COMO TROCAR A TELA, OU UTILIZAR PYRFC PARA COMMITAR E CRIAR DEIRETAMENTE O DIAGRAMA
        # ===================================================================================================

        network_creation(data)

        line += 1

    bot.alert(title='BotText', text='Programa encerrado!')