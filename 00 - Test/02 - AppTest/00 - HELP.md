PROCURAR INFORMAÇÕES DE TABELAS ESPECÍFICAS OU PERSONALIZADAS

result = conn.call(
        'RFC_READ_TABLE',
        QUERY_TABLE='Z22I0055_MD',
        ROWCOUNT=5,
        OPTIONS=[{
            'TEXT': f"QMNUM = '{notif_number}'"
        }],
        DELIMITER='|'
    )

---------------------------------------------------------------------------------------

FUNÇÃO PARA VISUALIZAR TODOS OS CAMPOS DA TABELA (SEUS RESPECTIVOS NOMES REAIS)

notif_result = conn.call('BAPI_ALM_NOTIF_GET_DETAIL', NUMBER=notif_number)
for key, value in notif_result.items():
    print(f"\n--- {key} ---")
    print(value)

OU

notif_result = conn.call('BAPI_ALM_NOTIF_GET_DETAIL', NUMBER=notif_number)
for key in notif_result.items():
    print(f"\n--- {key} ---")