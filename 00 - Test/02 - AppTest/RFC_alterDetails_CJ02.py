from pyrfc import Connection

sap_conn_params = {
    'user': 'MAO8CT',
    'passwd': '86IQ3J$.7vCj@',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

project_def = 'LP-050771'
nova_descricao = "Teste"

try:
    conn = Connection(**sap_conn_params)
    print('✅ Conectado ao SAP')

    # --- Atualizar cabeçalho do projeto (PROJ)
    result_proj = conn.call(
        'BAPI_PROJECTDEF_UPDATE',
        PROJECT_DEFINITION_STRU=project_def,
        PROJECT_DEFINITION_UPD={
            'PROJECT_DEFINITION': project_def,
            'DESCRIPTION': nova_descricao
        }
    )

    print("🟨 Resultado BAPI_PROJECTDEF_UPDATE:")
    print(result_proj)

    # --- Atualizar todos os elementos WBS vinculados ao projeto (PRPS)
    # Primeiro, buscar os elementos WBS do projeto
    estrutura = conn.call(
        'BAPI_PROJECT_STRUCTURE_GET',
        PROJECT_DEFINITION=project_def
    )

    wbs_elements = estrutura.get('WBS_ELEMENTS', [])

    # Prepara as alterações
    wbs_changes = []
    for wbs in wbs_elements:
        wbs_id = wbs['WBS_ELEMENT']
        wbs_changes.append({
            'WBS_ELEMENT': wbs_id,
            'DESCRIPTION': nova_descricao
        })

    # Envia alterações dos WBS (PRPS)
    if wbs_changes:
        result_prps = conn.call(
            'BAPI_BUS2054_CHANGE_MULTI',
            WBS_ELEMENT_UPD=wbs_changes
        )

        print("🟨 Resultado BAPI_BUS2054_CHANGE_MULTI:")
        print(result_prps)
    else:
        print("⚠️ Nenhum elemento WBS encontrado para o projeto.")

    # --- Commitar as alterações
    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print("✅ Alterações confirmadas com sucesso.")

except Exception as e:
    print(f"❌ Erro: {e}")