from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

notif_number = "14081856".zfill(12)

try:
    conn = Connection(**sap_conn_params)
    print("✅ Conectado ao SAP")

    result = conn.call(
        'RFC_READ_TABLE',
        QUERY_TABLE='Z22I0055_MD',
        ROWCOUNT=5,
        OPTIONS=[{
            'TEXT': f"QMNUM = '{notif_number}'"
        }],
        DELIMITER='|'
    )

    if result['DATA']:
        raw_line = result['DATA'][0]['WA']
        fields = raw_line.split('|')

        print('\n----- Informações Gerais -----')

        for i, field in enumerate(fields):
            print(f"Campo {i}: {field.strip()}")

        description = fields[5].strip()
        part_number = fields[6].strip()
        cost = fields[10].strip()
        date = fields[12].strip()

        print('\n----- Informações Importantes -----')
        print(f'Denominação: {description}\nPart Number: {part_number}\nCusto Orçado: {cost}\nPrazo Final: {date}\n')

except Exception as e:
    print(f"❌ Erro ao conectar ou chamar BAPI: {e}")