from pyrfc import Connection


# === 1. Conexão com o SAP ===
def conectar_sap(params):
    try:
        conn = Connection(**params)
        print("✅ Conectado ao SAP com sucesso!")
        return conn
    except Exception as e:
        raise Exception(f"Erro ao conectar ao SAP: {e}")


# === 2. Atualização do projeto ===
def atualizar_projeto(conn):
    i_project_definition = {
        'PROJECT_DEFINITION': 'LP-044055',
        'DESCRIPTION': 'Alteração automática via Python',
    }

    i_project_definition_upd = {
        'PROJECT_DEFINITION': 'X',
        'DESCRIPTION': 'X',
    }

    i_method_project = {
        'METHOD': 'UPDATE'
    }

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

    erro_encontrado = False
    if 'RETURN' in result:
        for msg in result['RETURN']:
            print(f"{msg['TYPE']}: {msg['MESSAGE']}")
            if msg['TYPE'] in ('E', 'A'):
                erro_encontrado = True

    if erro_encontrado:
        raise Exception("Erro na atualização do projeto.")


# === 3. Alteração do receptor de custos ===
def alterar_receptor_custos(conn, wbs_element, novo_receptor):
    # 3.1 - Obter OBJNR
    objnr_result = conn.call('RFC_READ_TABLE',
                             QUERY_TABLE='PRPS',
                             DELIMITER='|',
                             FIELDS=[{'FIELDNAME': 'OBJNR'}],
                             OPTIONS=[f"PSPNR = '{wbs_element}'"])
    if not objnr_result['DATA']:
        raise Exception("OBJNR não encontrado para o WBS.")

    objnr = objnr_result['DATA'][0]['WA'].split('|')[0].strip()
    print(f"OBJNR encontrado: {objnr}")

    # 3.2 - Ler regra de liquidação
    regra_atual = conn.call('K_SETTLEMENT_RULE_READ', OBJNR=objnr, OBJART='PROJ')
    settlement_rules = regra_atual.get('SETTL_RULE', [])
    if not settlement_rules:
        raise Exception("Nenhuma regra de liquidação encontrada.")

    for regra in settlement_rules:
        print(f"Regra antes: {regra}")
        regra['EMPGE'] = novo_receptor
        print(f"Regra depois: {regra}")

    # 3.3 - Validar regra
    conn.call('K_SETTLEMENT_RULE_CHECK', OBJNR=objnr, SETTL_RULE=settlement_rules)
    print("Regra validada com sucesso.")

    # 3.4 - Salvar regra
    conn.call('K_SETTLEMENT_RULE_SAVE', OBJNR=objnr, SETTL_RULE=settlement_rules)
    print("Regra salva com sucesso.")


# === 4. Commit ===
def confirmar_transacao(conn):
    conn.call('BAPI_TRANSACTION_COMMIT', WAIT='X')
    print("✅ Commit realizado com sucesso no SAP.")


# === 5. Rollback ===
def cancelar_transacao(conn):
    try:
        conn.call('BAPI_TRANSACTION_ROLLBACK')
        print("⚠️ Rollback executado no SAP.")
    except Exception as e:
        print(f"❗ Erro ao tentar fazer rollback: {e}")


# === 6. Main ===
def main():
    sap_connection_params = {
        'user': 'SEU_USUARIO',
        'passwd': 'SUA_SENHA',
        'ashost': 'rb3ps0a0.server.bosch.com',
        'sysnr': '00',
        'client': '011',
        'lang': 'PT'
    }

    conn = None
    try:
        conn = conectar_sap(sap_connection_params)
        atualizar_projeto(conn)
        alterar_receptor_custos(conn, 'LP-044055', 'RECEPTOR001')
        confirmar_transacao(conn)
        print("🎉 Processo finalizado com sucesso!")
    except Exception as e:
        print(f"❌ Erro: {e}")
        if conn:
            cancelar_transacao(conn)


if __name__ == '__main__':
    main()