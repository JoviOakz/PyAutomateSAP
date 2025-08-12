from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

# ----- VALORES ORIGINAIS -----
# Descrição: BATENTE
# Requerente: 68540003
# Data início: 15.07.2025 ou 20250715
# Data fim: 18.07.2025 ou 20250718

proj_num = 'LP-050771'
new_description = 'BATENTE Teste'
new_applicant = '68540028'
initial_date = '20250111'
final_date = '20251101'

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    proj_struct = {
        'PROJECT_DEFINITION': proj_num,
        'DESCRIPTION': new_description,
        'APPLICANT_NO': new_applicant,
        'START': initial_date,
        'FINISH': final_date
    }

    proj_struct_upd = {
        'DESCRIPTION': 'X',
        'APPLICANT_NO': 'X',
        'START': 'X',
        'FINISH': 'X'
    }

    proj_result = conn.call(
        'BAPI_PROJECTDEF_UPDATE',
        CURRENTEXTERNALPROJE=proj_num,
        PROJECT_DEFINITION_STRU=proj_struct,
        PROJECT_DEFINITION_UP=proj_struct_upd
    )

    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print('✅ Data updated successfully')

except Exception as e:
    print(f'Message: SAP connection or project update failed\nError: {e}')