from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

proj_num = 'LP-050771'
new_description = 'BATENTE Teste'

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    proj_struct = {
        'PROJECT_DEFINITION': proj_num,
        'DESCRIPTION': new_description
    }

    proj_struct_upd = {
        'DESCRIPTION': 'X'
    }

    result = conn.call(
        'BAPI_PROJECTDEF_UPDATE',
        CURRENTEXTERNALPROJE=proj_num,
        PROJECT_DEFINITION_STRU=proj_struct,
        PROJECT_DEFINITION_UP=proj_struct_upd
    )

    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print('✅ Data updated successfully')

except Exception as e:
    print(f'Message: SAP connection or project update failed\nError: {e}')