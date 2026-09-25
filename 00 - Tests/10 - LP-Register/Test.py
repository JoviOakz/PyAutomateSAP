# ===== LIBRARIES =====

import pyautogui as bot
import oracledb
import pandas as pd
import pyperclip as pc
from datetime import date, datetime
import time
import re

# ===== GLOBAL SETTINGS =====

INSTANT_CLIENT_PATH = r'C:\oracle\instantclient_23_0'

bot.FAILSAFE = True
bot.PAUSE = 0.85

# ===== INITIAL ACTION =====

bot.click(1802, 14)

# ===== STATIC FUNCTIONS =====

def press_key(key, times):
    for _ in range(times):
        if key == 'winr':
            bot.hotkey('win', 'r')
        elif key == 'ctrla':
            bot.hotkey('ctrl', 'a')
        elif key == 'ctrlc':
            bot.hotkey('ctrl', 'c')
        elif key == 'ctrlv':
            bot.hotkey('ctrl', 'v')
        elif key == 'ctrlf9':
            bot.hotkey('ctrl', 'f9')
        elif key == 'ctrltab':
            bot.hotkey('ctrl', 'tab')
        elif key == 'ctrlstab':
            bot.hotkey('ctrl', 'shift', 'tab')
        elif key == 'ctrlsf12':
            bot.hotkey('ctrl', 'shift', 'f12')
        elif key == 'stab':
            bot.hotkey('shift', 'tab')
        elif key == 'sf1':
            bot.hotkey('shift', 'f1')
        elif key == 'alte':
            bot.hotkey('alt', 'e')
        elif key == 'ctrls':
            bot.hotkey('ctrl', 's')
        else:
            bot.press(key)

def wait_event(img, region=None, timeout=10):
    inicio = time.time()

    while time.time() - inicio < timeout:
        try:
            local = bot.locateOnScreen(img, region=region, grayscale=True, confidence=0.9)

            if local:
                return local
        except:
            pass

        time.sleep(0.5)
    return None

# ===== FUNCTIONS =====

def sap_start():
    press_key('winr', 1)
    bot.typewrite('saplogon')
    press_key('enter', 1)

    if wait_event('images/SAP_1.png'):
        pass
    else:
        raise ValueError('|> SAP logon screen not found <|')

    press_key('stab', 1)
    bot.typewrite('ps0')
    press_key('enter', 1)

    if wait_event('images/SAP_2.png'):
        pass
    else:
        raise ValueError('|> SAP screen not found <|')

    bot.sleep(1)
    bot.typewrite('CJ02')
    press_key('enter', 1)

def close_sap():
    if wait_event('images/DIAGRAM_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Diagram screen not found <|')
    
    press_key('winr', 1)
    bot.typewrite('cmd /c taskkill /f /im saplogon.exe')
    press_key('enter', 1)

def wbs_element_creation(index, item):
    if wait_event('images/PROJECT_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Project screen not found <|')
    
    press_key('ctrla', 1)
    bot.typewrite(item[10])
    press_key('enter', 1)

    bot.PAUSE = 0.35
    
    if wait_event('images/PROJECT_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Project screen not found <|')

    press_key('ctrltab', 2)
    press_key('down', 1)
    press_key('ctrla', 1)
    press_key('ctrlc', 1)
    status = pc.paste().strip()

    if 'ABER' in status:
        pass
    elif 'LIB' in status:
        press_key('f3', 1)
        data[index].append('Liberado')
        return
    elif 'ENTE' or 'ENCE' in status:
        press_key('f3', 1)
        data[index].append('Encerrado')
        return

    press_key('stab', 1)
    press_key('ctrla', 1)
    part_number = re.sub(r'[-./POSpos& ]', '', item[5]).strip()
    pc.copy(item[6])

    bot.PAUSE = 0.85

    if part_number.isdigit():
        bot.typewrite(item[5] + ' - ')

    press_key('ctrlv', 1)
    press_key('ctrlf9', 1)

    if wait_event('images/WBS_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º WBS screen not found <|')

    bot.PAUSE = 0.35
    
    press_key('ctrltab', 4)
    press_key('tab', 1)
    press_key('ctrla', 1)

    keys = {
        'TEF': '68540012',
        'QMM': '68540007',
        'MFW1': '68540001',
        'MFE2': '68540002',
        'MFE3': '68540003',
    }

    iss_dept = None

    for key, value in keys.items():
        if key in item[8]:
            iss_dept = value
            break
        else:
            iss_dept = '68540028'

    bot.typewrite(iss_dept)
    press_key('ctrlstab', 3)
    press_key('right', 4)
    press_key('enter', 1)

    if wait_event('images/WBS_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º WBS screen not found <|')

    press_key('tab', 4)
    bot.typewrite(item[7])
    bot.sleep(0.5)
    press_key('tab', 2)
    bot.typewrite(str(int(float(item[4]))))
    press_key('tab', 1)
    bot.typewrite('PC')
    press_key('tab', 1)
    press_key('down', 1)
    bot.typewrite(str(int(float(item[9]))))
    press_key('tab', 1)
    bot.typewrite('BRL')
    press_key('enter', 1)
    bot.sleep(1.25)

    bot.PAUSE = 0.35

    press_key('stab', 4)
    press_key('enter', 1)

    bot.PAUSE = 0.85

    if wait_event('images/PARAMETERS_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Parameters screen not found <|')

    press_key('tab', 1)

    liquidation_object = item[2].strip().split('.')[0]
    alocation = '07'
    
    if liquidation_object.startswith('685') and len(liquidation_object) == 6:
        bot.typewrite('ZPS001')
        if len(item[3].strip()) == 1:
            alocation = '0' + item[3].strip()
        else:
            alocation = item[3].strip()
    elif liquidation_object.startswith('LP-'):
        bot.typewrite('ZPS007')
    elif liquidation_object.startswith('BM'):
        bot.typewrite('ZPS007')
    else:
        bot.typewrite('ZPS003')

    bot.sleep(1)
    press_key('tab', 1)
    bot.typewrite(alocation)
    bot.sleep(1)

    press_key('f3', 1)

    if wait_event('images/PARAMETERS_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Parameters screen not found <|')
    
    press_key('tab', 1)
    bot.typewrite(item[2].split('.')[0])
    press_key('f3', 1)

    if wait_event('images/WBS_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Return error <|')
    
    press_key('f3', 1)

    if wait_event('images/RETURN.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Return error <|')
    
    press_key('sf1', 1)

    if wait_event('images/PROJECT_3.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 3º WBS screen not found <|')
    
    part_number = re.sub(r'[-./POSpos& ]', '', item[5]).strip()
    pc.copy(item[6])

    if part_number.isdigit():
        bot.typewrite(item[5] + ' - ')

    press_key('ctrlv', 1)

    bot.PAUSE = 0.35

    press_key('ctrltab', 4)
    press_key('tab', 1)
    bot.typewrite(iss_dept)
    bot.sleep(1)
    press_key('down', 1)
    bot.typewrite(date.today().strftime('%d.%m.%Y'))
    bot.sleep(1)
    press_key('down', 1)
    bot.typewrite(date.today().strftime('%d.%m.%Y'))
    bot.sleep(1)
    press_key('alte', 1)
    press_key('s', 1)
    press_key('i', 1)
    bot.sleep(1.5)

    bot.PAUSE = 0.85

    data[index].append('Cadastrado parcial')
    press_key('ctrls', 1)
    
def cn21_config():
    if wait_event('images/PROJECT_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> Project screen not found <|')
    
    press_key('ctrlstab', 1)
    press_key('tab', 1)
    bot.typewrite('/ncn21')
    press_key('enter', 1)

    if wait_event('images/DIAGRAM_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Diagram screen not found <|')
    
    bot.PAUSE = 0.15

    press_key('right', 1)
    press_key('left', 3)
    press_key('right', 2)
    
    bot.PAUSE = 0.85

    press_key('tab', 1)
    bot.typewrite('BP01')
    press_key('tab', 1)
    bot.typewrite('6854')
    bot.sleep(1.15)

    bot.PAUSE = 0.35

    press_key('tab', 1)

    mrp = item[1].strip()

    if mrp == 'Yesica Gonzalez':
        bot.typewrite('I33')
    elif mrp == 'Rodrigo Melo':
        bot.typewrite('I49')
    elif mrp == 'Marcelo Simoes':
        bot.typewrite('I55')
    elif mrp == 'Edson Bento':
        bot.typewrite('I38')
    elif mrp == 'Thais Fischer':
        bot.typewrite('I55')
    elif mrp == 'Joao Franca':
        bot.typewrite('I31')
    else:
        bot.typewrite('I39')

    press_key('stab', 3)
    bot.PAUSE = 0.85

def mrp_config():
    mrp = item[1].strip()
    resp_change = False

    if index > 0:
        previous_mrp = data[index - 1][1].strip()

        if mrp != previous_mrp:
            resp_change = True

    if resp_change:
        press_key('tab', 3)
        bot.sleep(1.25)

        if mrp == 'Yesica Gonzalez':
            bot.typewrite('I33')
        elif mrp == 'Rodrigo Melo':
            bot.typewrite('I49')
        elif mrp == 'Marcelo Simoes':
            bot.typewrite('I55')
        elif mrp == 'Edson Bento':
            bot.typewrite('I38')
        elif mrp == 'Thais Fischer':
            bot.typewrite('I55')
        elif mrp == 'Joao Franca':
            bot.typewrite('I31')
        else:
            bot.typewrite('I39')

        bot.sleep(0.5)

def diagram_creation(index, item):
    bot.PAUSE = 0.35

    if wait_event('images/DIAGRAM_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Diagram screen not found <|')

    mrp_config()
    
    bot.PAUSE = 0.85
    
    press_key('enter', 1)

    if wait_event('images/VALUE.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> Value box not found <|')

    bot.typewrite(item[10].replace('-', ''))
    press_key('enter', 1)

    if wait_event('images/DIAGRAM_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Diagram screen not found <|')

    press_key('ctrltab', 1)
    press_key('right', 1)
    press_key('enter', 1)

    if wait_event('images/ATTRIBUITION_1.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 1º Attribuition screen not found <|')

    press_key('tab', 2)
    bot.typewrite(item[10].replace('-', ''))
    bot.sleep(1.15)

    bot.PAUSE = 0.35

    press_key('stab', 2)
    press_key('right', 3)
    press_key('enter', 1)

    if wait_event('images/ATTRIBUITION_2.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Attribuition screen not found <|')

    press_key('tab', 1)
    press_key('enter', 1)

    if wait_event('images/ATTRIBUITION_3.png'):
        pass
    else:
        data[index].append('Error')
        save_excel(data)
        raise ValueError('|> 2º Attribuition screen not found <|')

    press_key('ctrltab', 2)
    press_key('ctrla', 1)

    heijunka = ''
    description = ''
    responsible = ''

    part_number = re.sub(r'[-./POSpos& ]', '', item[5]).strip()

    if item[1] == 'Yesica Gonzalez' or item[1] == 'Rodrigo Melo':
        heijunka = 'HEIJUNKA\n'

    if part_number.isdigit():
        description = item[5] + ' - '

    description += item[6]
    responsible = '\nResp. ' + item[1]
    full_text = heijunka + description + responsible

    bot.typewrite(full_text)
    bot.sleep(1.15)
    press_key('ctrlsf12', 1)

    if wait_event('images/ATTRIBUITION_4.png'):
        pass
    else:
        if wait_event('images/ATTRIBUITION_5.png'):
            press_key('f12', 1)
            bot.sleep(1.15)
            press_key('f12', 1)
            bot.sleep(1.15)
            press_key('tab', 1)
            press_key('enter', 1)
            data[index].append('Erro no cadastro do diagrama')
            return

    bot.PAUSE = 0.85

    bot.sleep(0.5)
    data[index].append('Cadastrado')
    press_key('ctrls', 1)
    bot.sleep(1.5)

def save_excel(data, active_lps):
    record_time = datetime.now().strftime('%d-%m_%H-%M')

    df = pd.DataFrame(data)

    EXCEL_PATH = f'./Record_{record_time}.xlsx'
    df.to_excel(
        EXCEL_PATH,
        engine='openpyxl',
        index=False
    )

    if active_lps:
        df = pd.DataFrame(active_lps)

        EXCEL_PATH = f'./Active_LPs_{record_time}.xlsx'
        df.to_excel(
            EXCEL_PATH,
            engine='openpyxl',
            index=False
        )

# ===== PROGRAM CONFIGURATION =====

try:
    oracledb.init_oracle_client(lib_dir=INSTANT_CLIENT_PATH)
    
except Exception as e:
    print(f'Failed to initialize Oracle client: {e}')

USER = 'MAO8CT'
PASS = '49l1)f=f3q6A'
dsn = 'REDLake_ZeusP_Consumer_Common.world'

try:
    with oracledb.connect(user=USER, password=PASS, dsn=dsn) as connection:
        with connection.cursor() as conn:
            query = f'''
                SELECT DISTINCT
                    Z54.TIPO_DEMANDA,
                    Z54.NAME_LIST_PLANEJADOR AS RESPONSAVEL,
                    Z54.NUM_RS               AS OBJ_LIQUIDACAO,
                    Z54.ABSCH                AS ESQ_ALOCACAO,
                    Z55.MENGE                AS QUANTIDADE,
                    Z55.NR_TIPO_PARTNR       AS PARTNUMBER,
                    Z55.POST1                AS DENOMINACAO_ITEM,
                    Z55.ENTREGAR_A           AS ENTREGAR_A,
                    Z54.DEPARTMENT_EMIT      AS DEPT_EMIT,
                    Z55.ESTIMATED_COSTS      AS CUSTO_ESTIMADO,
                    PROJ.PSPID_EDIT          AS LP
                FROM MARD_MDNA.V_CUSN_Z22I0055_MD_B2 Z55
                LEFT JOIN MARD_MDNA.V_CUSN_Z22I0054_MD_B2 Z54
                    ON Z54.QMNUM = Z55.QMNUM
                LEFT JOIN MARD_MDNA.V_CUSN_PROJ_B2 PROJ
                    ON PROJ.PSPNR = Z55.PSPNR
                WHERE Z54.PARNR_PLANEJADOR IN ('IOS3CT','ENB9CT','MEO9CT','LIY1CT','FIH9CT','FRJ1CT','NUR3CT','LRI2CT','MER7CT','COH1CT','ADB2CT')
                    AND Z54.TECH_TIMESTAMP >= TIMESTAMP '2026-01-01 00:00:00'
                    AND Z54.TECH_TIMESTAMP <  TIMESTAMP '2027-01-01 00:00:00'
                    AND PROJ.PSPID_EDIT IS NOT NULL
                    AND PROJ.AEDAT = '00000000'
                ORDER BY PROJ.PSPID_EDIT ASC
            '''

            conn.execute(query)
            data = conn.fetchall()
            data = [list(item) for item in data]
            active_lps = [linha for linha in data if linha[0] == 'A']
            data = [linha for linha in data if linha[0] != 'A']

except oracledb.Error as e:
    raise ValueError('Connection failed: {e}')

# ===== MAIN =====

if __name__ == '__main__':
    sap_start()

    for index, item in enumerate(data):
        wbs_element_creation(index, item)

    cn21_config()

    for index, item in enumerate(data):
        if item[11] == 'Cadastrado parcial':
            diagram_creation(index, item)

    save_excel(data, active_lps)
    close_sap()