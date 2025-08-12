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

    # try:
    #     # PARAMETERS: FUNCTION MODULE
    #     func_desc = conn.get_function_description('BAPI_PROJECT_MAINTAIN')

    #     print('\n📌 BAPI_PROJECT_MAINTAIN - Function Module:\n')
    #     for param in func_desc.parameters:
    #         print(param)

    # except Exception as e:
    #     print(f'Message: error with Function Modules parameters\nError: {e}')

    try:
        # PARAMETERS: FIELDS
        func_desc = conn.get_function_description('BAPI_PROJECT_MAINTAIN')

        for param in func_desc.parameters:
            if param['name'] == 'I_METHOD_PROJECT':
                print('\n\n📌 I_METHOD_PROJECT - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Table parameters\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')