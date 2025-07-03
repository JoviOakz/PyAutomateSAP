from pyrfc import Connection

def criar_diagrama_rede_rollback():
    sap_connection_params = {
        'user': 'ARA2CT',
        'passwd': 'AKDAgmr007@@y',
        'ashost': 'rb3ps0a0.server.bosch.com',
        'sysnr': '00',
        'client': '011',
        'lang': 'PT'
    }

    try:
        conn = Connection(**sap_connection_params)
        print("✅ Conectado ao SAP com sucesso!")

        wbs_element = 'LP-050562'
        centro = '6854'
        planejador = 'I55'
        tipo_rede = 'BP01'
        perfil = 'ZBP0001'

        i_project_definition = {
            'PS_PSP_PNR': wbs_element,
            'PS_PSP_PRO': wbs_element,
            'PROFIDNZPL': perfil,
            'WERKS_D': centro,
            'CO_DISPO': planejador,
            'PS_AUFART': tipo_rede,
            'PVARFIELD': wbs_element,
        }

        i_project_definition_upd = {
            'PS_PSP_PNR': wbs_element,
            'PS_PSP_PRO': wbs_element,
            'PROFIDNZPL': perfil,
            'WERKS_D': centro,
            'CO_DISPO': planejador,
            'PS_AUFART': tipo_rede,
            'PVARFIELD': wbs_element,
        }

        i_network = []
        i_network_update = []
        i_method_project = []

        result = conn.call('BAPI_PROJECT_MAINTAIN',
                           I_NETWORK=i_network,
                           I_NETWORK_UPDATE=i_network_update,
                           I_METHOD_PROJECT=i_method_project,
                           I_PROJECT_DEFINITION=i_project_definition,
                           I_PROJECT_DEFINITION_UPD=i_project_definition_upd
                          )

        for msg in result.get('RETURN', []):
            print(f"{msg['TYPE']}: {msg['MESSAGE']}")

        if any(msg['TYPE'] in ('E', 'A') for msg in result.get('RETURN', [])):
            conn.call('BAPI_TRANSACTION_ROLLBACK')
            print("❌ Erro detectado. Transação revertida (rollback).")
        else:
            conn.call('BAPI_TRANSACTION_COMMIT')
            print("✅ Transação comitada com sucesso!")

    except Exception as e:
        print("❌ Erro na conexão ou execução da BAPI:", e)

if __name__ == "__main__":
    criar_diagrama_rede_rollback()