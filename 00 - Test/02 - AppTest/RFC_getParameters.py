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
        func_desc = conn.get_function_description('MAP_BAPI_WBS_ELEMENT_2_PRPS')

        print('\n📌 MAP_BAPI_WBS_ELEMENT_2_PRPS - Parameters:\n')
        for param in func_desc.parameters:
            print(param)

    except Exception as e:
        print(f'Message: error with Function Modules parameters\nError: {e}')

    try:
        # PARAMETERS: TABLE
        func_desc = conn.get_function_description('MAP_BAPI_WBS_ELEMENT_2_PRPS')

        for param in func_desc.parameters:
            if param['name'] == 'BAPI_WBS_ELEMENT':
                print('\n\n📌 BAPI_WBS_ELEMENT - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Table parameters\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')