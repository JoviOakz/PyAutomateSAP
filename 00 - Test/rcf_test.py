from pyrfc import Connection

# === Conectar ao SAP ===
sap_connection_params = {
    'user': 'SEU_USUARIO',
    'passwd': 'SUA_SENHA',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

try:
    conn = Connection(**sap_connection_params)
    print("Conectado ao SAP com sucesso!")
except Exception as e:
    print(f"Erro na conexão com SAP: {e}")
    exit()

try:
    # === Atualização do Projeto ===
    i_project_definition = {
        'PROJECT_DEFINITION': 'LP-044055',
        'DESCRIPTION': 'Alteração automática via Python',
    }

    i_project_definition_upd = {
        'PROJECT_DEFINITION': 'X',
        'DESCRIPTION': 'X',
    }

    i_method_project = {'METHOD': 'UPDATE'}

    i_wbs_element_table_update = [{
        'WBS_ELEMENT': 'LP-044055',
        'APPLICANT_NO': '68540012'
    }]

    i_wbs_element_table_update_upd = [{
        'WBS_ELEMENT': ' ',
        'APPLICANT_NO': 'X'
    }]

    result = conn.call('BAPI_PROJECT_MAINTAIN',
        I_PROJECT_DEFINITION=i_project_definition,
        I_PROJECT_DEFINITION_UPD=i_project_definition_upd,
        I_METHOD_PROJECT=i_method_project,
        I_WBS_ELEMENT_TABLE_UPDATE=i_wbs_element_table_update,
        I_WBS_ELEMENT_TABLE_UPDATE_UPD=i_wbs_element_table_update_upd
    )

    for message in result.get('RETURN', []):
        print(f"{message['TYPE']}: {message['MESSAGE']}")
        if message['TYPE'] in ['E', 'A']:  # Erro ou Abort
            raise Exception(f"Erro na manutenção do projeto: {message['MESSAGE']}")

    # === Consulta do OBJNR ===
    wbs_element = 'LP-044055'
    novo_receptor_custos = 'RECEPTOR001'

    objnr_result = conn.call('RFC_READ_TABLE',
                             QUERY_TABLE='PRPS',
                             DELIMITER='|',
                             FIELDS=[{'FIELDNAME': 'OBJNR'}],
                             OPTIONS=[f"PSPNR = '{wbs_element}'"])
    if not objnr_result['DATA']:
        raise Exception("OBJNR não encontrado para o WBS.")

    objnr = objnr_result['DATA'][0]['WA'].split('|')[0].strip()

    # === Leitura e alteração da regra de liquidação ===
    regra_atual = conn.call('K_SETTLEMENT_RULE_READ',
                            OBJNR=objnr,
                            OBJART='PROJ')

    settlement_rules = regra_atual.get('SETTL_RULE', [])
    if not settlement_rules:
        raise Exception("Nenhuma regra de liquidação encontrada.")

    for regra in settlement_rules:
        regra['EMPGE'] = novo_receptor_custos

    # === Validação ===
    validacao = conn.call('K_SETTLEMENT_RULE_CHECK',
                          OBJNR=objnr,
                          SETTL_RULE=settlement_rules)

    for message in validacao.get('RETURN', []):
        if message['TYPE'] in ['E', 'A']:
            raise Exception(f"Erro na validação: {message['MESSAGE']}")

    # === Salvando regra ===
    result_save = conn.call('K_SETTLEMENT_RULE_SAVE',
                            OBJNR=objnr,
                            SETTL_RULE=settlement_rules)

    for message in result_save.get('RETURN', []):
        if message['TYPE'] in ['E', 'A']:
            raise Exception(f"Erro ao salvar regra: {message['MESSAGE']}")

    # === Commit final ===
    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print("✅ Todas as alterações foram aplicadas com sucesso no SAP.")

except Exception as error:
    print(f"❌ Ocorreu um erro durante o processo: {error}")
    print("Executando rollback...")
    try:
        conn.call('BAPI_TRANSACTION_ROLLBACK')
        print("Rollback executado com sucesso.")
    except Exception as rollback_error:
        print(f"Erro ao executar rollback: {rollback_error}")