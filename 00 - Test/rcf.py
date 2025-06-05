from pyrfc import Connection

# Dados de conexão com SAP
sap_connection_params = {
    'user': 'SEU_USUARIO',
    'passwd': 'SUA_SENHA',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',  # Número do sistema
    'client': '011',  # Cliente SAP
    'lang': 'PT'
}

# Conectar
conn = Connection(**sap_connection_params)

print("Conectado ao SAP com sucesso!")

# Parâmetros de exemplo
project_data = {
    'PROJECT_DEFINITION': 'PROJ001',
    'WBS_ELEMENT': 'PROJ001-01'
}

result = conn.call('BAPI_PROJECT_MAINTAIN', PROJECTSTRUCTURE=[project_data])

print(result)

network_data = {
    'NETWORK': 'NET001',
    'DESCRIPTION': 'Nova rede de projeto',
    # Outros campos necessários...
}

result = conn.call('BAPI_NETWORK_CREATE', NETWORKHEADER=network_data)

print(result)

if 'RETURN' in result:
    for message in result['RETURN']:
        print(f"{message['TYPE']}: {message['MESSAGE']}")