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
nova_descricao = "Retrabalho Holder"

try:
    conn = Connection(**sap_conn_params)
    print("✅ Conectado ao SAP")

    # Monta estrutura de dados a atualizar
    proj_stru = {
        "PROJECT_DEFINITION": project_def,  # Código do projeto
        "DESCRIPTION": nova_descricao             # Descrição do projeto (tabela PROJ)
    }

    # Estrutura de atualização (indica quais campos serão alterados)
    proj_up = {
        "DESCRIPTION": "X"  # Marca que a descrição será atualizada
    }

    # Chamada da BAPI
    result = conn.call(
        "BAPI_PROJECTDEF_UPDATE",
        CURRENTEXTERNALPROJE=project_def,     # Identificação do projeto
        PROJECT_DEFINITION_STRU=proj_stru,    # Novos dados
        PROJECT_DEFINITION_UP=proj_up         # Campos a alterar
    )

    # Commit das alterações
    conn.call("BAPI_TRANSACTION_COMMIT", WAIT="X")
    print("✅ Alteração confirmada com sucesso.")

except Exception as e:
    print(f"❌ Erro: {e}")