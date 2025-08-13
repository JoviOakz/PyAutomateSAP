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
new_description = 'BATENTE 2'
new_applicant = '68540028'
initial_date = '20250111'
final_date = '20251101'

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    # ----- Update project -----
    try:
        proj_struct = {
            'PROJECT_DEFINITION': proj_num,
            'DESCRIPTION': new_description,
            # 'APPLICANT_NO': new_applicant,
            # 'START': initial_date,
            # 'FINISH': final_date
        }

        proj_struct_upd = {
            'DESCRIPTION': 'X',
            # 'APPLICANT_NO': 'X',
            # 'START': 'X',
            # 'FINISH': 'X'
        }

        proj_result = conn.call(
            'BAPI_PROJECTDEF_UPDATE',
            CURRENTEXTERNALPROJE=proj_num,
            PROJECT_DEFINITION_STRU=proj_struct,
            PROJECT_DEFINITION_UP=proj_struct_upd
        )

        print('✅ Project updated successfully')

    except Exception as e:
        print(f'Message: Project update failed\nError: {e}')

    # ----- Update WBS element -----
    try:
        wbs_struct = [{
            'WBS_ELEMENT': proj_num,
            'DESCRIPTION': new_description
        }]

        wbs_struct_upd = [{
            'DESCRIPTION': 'X'
        }]

        wbs_result = conn.call(
            'BAPI_BUS2054_CHANGE_MULTI',
            I_PROJECT_DEFINITION=proj_num,
            IT_WBS_ELEMENT=wbs_struct,
            IT_UPDATE_WBS_ELEMENT=wbs_struct_upd
        )

        print('✅ WBS element updated successfully')

    except Exception as e:
        print(f'Message: WBS element update failed\nError: {e}')

    # ----- Commit changes -----
    try:
        conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
        print('✅ LP registered successfully')

    except Exception as e:
        print(f'Message: Commit failed (check LP)\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection failed\nError: {e}')