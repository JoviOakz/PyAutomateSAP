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
    print("✅ Conectado ao SAP")

    # Obter a descrição da BAPI
    func_desc = conn.get_function_description("BAPI_PROJECTDEF_UPDATE")

    print("\n📌 Parâmetros da BAPI_PROJECTDEF_UPDATE:\n")
    for param in func_desc.parameters:  # Agora usando atributo parameters
        print(param)

except Exception as e:
    print(f"❌ Erro: {e}")

# =======================================================================================

try:
    conn = Connection(**sap_conn_params)
    print("✅ Conectado ao SAP")

    func_desc = conn.get_function_description("BAPI_PROJECTDEF_UPDATE")

    for param in func_desc.parameters:
        if param["name"] == "PROJECT_DEFINITION_STRU":
            print("\n📌 Campos de PROJECT_DEFINITION_STRU:")
            for field in param["type_description"].fields:
                print(field)  # Mostra o dicionário completo para sabermos o nome certo

except Exception as e:
    print(f"❌ Erro: {e}")