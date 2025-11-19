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
        func_desc = conn.get_function_description('BAPI_NETWORK_MAINTAIN')

        print('\n\n📌 BAPI_NETWORK_MAINTAIN - Function Module:\n')
        for param in func_desc.parameters:
            print(param)

    except Exception as e:
        print(f'Message: error with Function Modules parameters\nError: {e}')

    # ----- Parameters: Fields -----
    try:
        func_desc = conn.get_function_description('BAPI_NETWORK_MAINTAIN')

        for param in func_desc.parameters:
            if param['name'] == 'I_NETWORK':
                print('\n\n📌 I_NETWORK - Fields:\n')
                for field in param['type_description'].fields:
                    print(field)

    except Exception as e:
        print(f'Message: error with Fields parameters\nError: {e}')

    # ----- Table: Columns -----
    try:
        result = conn.call(
            'RFC_READ_TABLE',
            QUERY_TABLE='DD03L',
            DELIMITER='|',
            OPTIONS=[
                {'TEXT': "TABNAME = 'Z22I0055_MD'"},
                {'TEXT': "AND AS4LOCAL = 'A'"}
            ],
            FIELDS=[
                {'FIELDNAME': 'FIELDNAME'}
            ]
        )

        for row in result['DATA']:
            cols = row['WA'].split('|')
            print(cols)

    except Exception as e:
        print(f'Message: error with Columns parameters\nError: {e}')

except Exception as e:
    print(f'Message: SAP connection attempt failed\nError: {e}')