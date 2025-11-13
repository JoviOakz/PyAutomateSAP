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

def press_key(key, times):
    for _ in range(times):
        if key == 'ctrls':
            bot.hotkey('ctrl', 's')
        elif key == 'ctrlf9':
            bot.hotkey('ctrl', 'f9')
        else:
            bot.press(key)

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

        bot.typewrite(str(data['project']))
        bot.sleep(2)
        press_key('enter')
        bot.sleep(2)

        # essa verificação do tamanho pode ser feito uma unica vez e ser criado uma var para o nome "novo"
        # if (desc < 12c == norm+desc):
        #     bot.typewrite(norm+desc)

        press_key('tab', 3)

        # if (desc < 12c == norm+desc):
        #     bot.typewrite(norm+desc)

        press_key('ctrlf9', 1)







        

        line += 1

    bot.alert(title='BotText', text='Programa encerrado!')