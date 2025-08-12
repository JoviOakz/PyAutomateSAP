from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

notif_number = '14088187'.zfill(12)

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    try:
        result_0054 = conn.call(
            'RFC_READ_TABLE',
            QUERY_TABLE='Z22I0054_MD',
            ROWCOUNT=5,
            OPTIONS=[{
                'TEXT': f"QMNUM = '{notif_number}'"
            }],
            DELIMITER='|'
        )
     
        if result_0054['DATA']:
            raw_line = result_0054['DATA'][0]['WA']
            fields = raw_line.split('|')

            payee = fields[7].strip()
            resp = fields[9].strip()
            liquidation_obj = fields[12].strip()

            print('\n----- Informações da RS -----')
            print(f'Emitente: {payee}\nResponsável: {resp}\nObjeto de liquidação: {liquidation_obj}')
   
    except Exception as e:
        print(f'Message: error with Z22I0054_MD table\nError: {e}')

    try:
        result_0055 = conn.call(
            'RFC_READ_TABLE',
            QUERY_TABLE='Z22I0055_MD',
            ROWCOUNT=5,
            OPTIONS=[{
                'TEXT': f"QMNUM = '{notif_number}'"
            }],
            DELIMITER='|'
        )
     
        if result_0055['DATA']:
            raw_line = result_0055['DATA'][0]['WA']
            fields = raw_line.split('|')

            quantity = int(float(fields[4].strip()))
            description = fields[5].strip()
            part_number = fields[6].strip()
            payee = fields[9].strip()
            cost = fields[10].strip()
            date = fields[12].strip()

            date = str(date[6:8]) + '.' + str(date[4:6]) + '.' + str(date[:4])

            print(f'Quantidade: {quantity}\nDescrição: {description}\nNorma: {part_number}\nEntregar a: {payee}\nCusto orçado: {cost}\nPrazo final: {date}\n')
   
    except Exception as e:
        print(f'Message: error with Z22I0055_MD table\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')