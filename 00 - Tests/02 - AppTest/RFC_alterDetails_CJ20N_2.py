from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

project_number = 'LP-050771'
new_description = 'BATENTE Teste'

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    proj_struct = {
        'PROJECT_DEFINITION': project_number,
        'DESCRIPTION': new_description
    }

    proj_struct_upd = {
        'DESCRIPTION': 'X'
    }

    wbs_struct = [{
        'WBS_ELEMENT': project_number,
        'PROJECT_DEFINITION': project_number,
        'DESCRIPTION': new_description
    }]

    wbs_struct_upd = [{
        'DESCRIPTION': 'X'
    }]

    method_project = [
        {
            'REFNUMBER': '000001',
            'OBJECTTYPE': 'PROJECT',
            'METHOD': 'UPDATE',
            'OBJECTKEY': project_number
        },
        {
            'REFNUMBER': '000002',
            'OBJECTTYPE': 'WBS-ELEMENT',
            'METHOD': 'UPDATE',
            'OBJECTKEY': project_number
        }
    ]

    result = conn.call(
        'BAPI_PROJECT_MAINTAIN',
        I_PROJECT_DEFINITION=proj_struct,
        I_PROJECT_DEFINITION_UPD=proj_struct_upd,
        I_WBS_ELEMENT_TABLE=wbs_struct,
        I_WBS_ELEMENT_TABLE_UPDATE=wbs_struct_upd,
        I_METHOD_PROJECT=method_project
    )

    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print('✅ Data updated successfully')

except Exception as e:
    print(f'Message: SAP connection or project update failed\nError: {e}')