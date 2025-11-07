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

    # ----- Parameters: Function module -----
    try:
        func_desc = conn.get_function_description('BAPI_BUS2054_CHANGE_MULTI')

        print('\n📌 BAPI_BUS2054_CHANGE_MULTI - Function Module:\n')
        for param in func_desc.parameters:
            print(param)

    except Exception as e:
        print(f'Message: error with Function Modules parameters\nError: {e}')

    # ----- Parameters: Fields -----
    try:
        func_desc = conn.get_function_description('BAPI_BUS2054_CHANGE_MULTI')

        for param in func_desc.parameters:
            if param['name'] == 'I_PROJECT_DEFINITION':
                print('\n\n📌 I_PROJECT_DEFINITION - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Fields parameters\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')