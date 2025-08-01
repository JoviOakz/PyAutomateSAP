from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj',
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
 
    notif_result = conn.call('BAPI_ALM_NOTIF_GET_DETAIL', NUMBER=notif_number)
    
    for key, value in notif_result.items():
        print(f"\n--- {key} ---")
        print(value)

except Exception as e:
    print(f"❌ Erro ao conectar ou chamar BAPI: {e}")