from pyrfc import Connection

# === 1. Dados de conexão com SAP ===
sap_connection_params = {
    'user': 'SEU_USUARIO',
    'passwd': 'SUA_SENHA',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

# Conectar
conn = Connection(**sap_connection_params)
print("Conectado ao SAP com sucesso!")

# === 2. Atualização do Projeto (descrição, applicant) ===
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

# Chamada da BAPI para manutenção do projeto
result = conn.call('BAPI_PROJECT_MAINTAIN',
    I_PROJECT_DEFINITION=i_project_definition,
    I_PROJECT_DEFINITION_UPD=i_project_definition_upd,
    I_METHOD_PROJECT=i_method_project,
    I_WBS_ELEMENT_TABLE_UPDATE=i_wbs_element_table_update,
    I_WBS_ELEMENT_TABLE_UPDATE_UPD=i_wbs_element_table_update_upd
)

print("Resultado da manutenção do projeto:")
if 'RETURN' in result:
    for message in result['RETURN']:
        print(f"{message['TYPE']}: {message['MESSAGE']}")

# === 3. Alteração do Receptor de Apropriação de Custos (EMPGE) ===

# Dados necessários
wbs_element = 'LP-044055'
novo_receptor_custos = 'RECEPTOR001'  # <-- ajuste conforme necessário

# === 3.1 Consulta do OBJNR correspondente ao WBS ===
try:
    objnr_result = conn.call('RFC_READ_TABLE',
                             QUERY_TABLE='PRPS',
                             DELIMITER='|',
                             FIELDS=[{'FIELDNAME': 'OBJNR'}],
                             OPTIONS=[f"PSPNR = '{wbs_element}'"])
    if not objnr_result['DATA']:
        print("OBJNR não encontrado para o WBS.")
        exit()

    objnr = objnr_result['DATA'][0]['WA'].split('|')[0].strip()
    print(f"OBJNR encontrado: {objnr}")

except Exception as e:
    print(f"Erro ao consultar OBJNR: {e}")
    exit()

# === 3.2 Ler a regra de liquidação ===
try:
    regra_atual = conn.call('K_SETTLEMENT_RULE_READ',
                            OBJNR=objnr,
                            OBJART='PROJ')  # ou 'PEP', conforme sua configuração
    print("Regra de liquidação lida com sucesso.")
except Exception as e:
    print(f"Erro ao ler a regra de liquidação: {e}")
    exit()

# === 3.3 Alterar o campo EMPGE ===
try:
    settlement_rules = regra_atual.get('SETTL_RULE', [])
    if not settlement_rules:
        print("Nenhuma regra de liquidação encontrada.")
        exit()

    for regra in settlement_rules:
        print(f"Regra antes da alteração: {regra}")
        regra['EMPGE'] = novo_receptor_custos
        print(f"Regra após alteração: {regra}")

except Exception as e:
    print(f"Erro ao alterar a regra: {e}")
    exit()

# === 3.4 Validar a regra alterada ===
try:
    validacao = conn.call('K_SETTLEMENT_RULE_CHECK',
                          OBJNR=objnr,
                          SETTL_RULE=settlement_rules)
    print("Regra validada com sucesso.")
except Exception as e:
    print(f"Erro na validação da regra: {e}")
    exit()

# === 3.5 Gravar a regra alterada ===
try:
    result_save = conn.call('K_SETTLEMENT_RULE_SAVE',
                            OBJNR=objnr,
                            SETTL_RULE=settlement_rules)
    print("Regra de liquidação salva com sucesso.")
except Exception as e:
    print(f"Erro ao salvar a regra: {e}")
    exit()

print("Processo concluído com sucesso!")