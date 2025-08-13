from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

try:
    conn = Connection(**sap_conn_params)
    print('✅ SAP connected successfully')

    try:
        # PARAMETERS: FUNCTION MODULE
        func_desc = conn.get_function_description('BAPI_PROJECTDEF_UPDATE')

        print('\n📌 BAPI_PROJECTDEF_UPDATE - Function Module:\n')
        for param in func_desc.parameters:
            print(param)

    except Exception as e:
        print(f'Message: error with Function Modules parameters\nError: {e}')

    try:
        # PARAMETERS: FIELDS
        func_desc = conn.get_function_description('BAPI_PROJECTDEF_UPDATE')

        for param in func_desc.parameters:
            if param['name'] == 'PROJECT_DEFINITION_STRU':
                print('\n\n📌 PROJECT_DEFINITION_STRU - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Fields parameters\nError: {e}')

    try:
        # PARAMETERS: FIELDS 2
        func_desc = conn.get_function_description('BAPI_PROJECTDEF_UPDATE')

        for param in func_desc.parameters:
            if param['name'] == 'PROJECT_DEFINITION_UP':
                print('\n\n📌 PROJECT_DEFINITION_UP - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Fields parameters\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')