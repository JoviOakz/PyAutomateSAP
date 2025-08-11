from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

# =======================================================================================
# GET PARAMETERS -> FUNCTION MODULE

try:
    conn = Connection(**sap_conn_params)
    print("✅ Conectado ao SAP")

    # Obter a descrição da BAPI
    func_desc = conn.get_function_description("MAP_BAPI_WBS_ELEMENT_2_PRPS")

    print("\n📌 Parâmetros da MAP_BAPI_WBS_ELEMENT_2_PRPS:\n")
    for param in func_desc.parameters:  # Agora usando atributo parameters
        print(param)

except Exception as e:
    print(f"❌ Erro: {e}")

# =======================================================================================
# GET PARAMETERS -> TABLE

try:
    conn = Connection(**sap_conn_params)
    print("\n\n✅ Conectado ao SAP")

    func_desc = conn.get_function_description("MAP_BAPI_WBS_ELEMENT_2_PRPS")

    for param in func_desc.parameters:
        if param["name"] == "BAPI_WBS_ELEMENT":
            print("\n📌 Campos de BAPI_WBS_ELEMENT:")
            for field in param["type_description"].fields:
                print(field)  # Mostra o dicionário completo para sabermos o nome certo

except Exception as e:
    print(f"❌ Erro: {e}")