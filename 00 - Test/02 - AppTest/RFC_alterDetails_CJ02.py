from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

project_def = 'LP-050771'
new_description = 'BATENTE Teste'

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    proj_stru = {
        'PROJECT_DEFINITION': project_def,
        'DESCRIPTION': new_description
    }

    proj_up = {
        'DESCRIPTION': 'X'
    }

    result = conn.call(
        'BAPI_PROJECTDEF_UPDATE',
        CURRENTEXTERNALPROJE=project_def,
        PROJECT_DEFINITION_STRU=proj_stru,
        PROJECT_DEFINITION_UP=proj_up
    )

    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print('✅ Data updated successfully')

except Exception as e:
    print(f'Message: Connection or project update failed\nError: {e}')