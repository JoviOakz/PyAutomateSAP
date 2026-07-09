import oracledb

try:
    pasta_instant_client = r'C:\Oracle\instantclient_23_0' 
    
    oracledb.init_oracle_client(
        lib_dir=pasta_instant_client,
        config_dir=r'C:\Oracle\ora_net'
    )

    print('🚀 Modo Thick ativado com sucesso!')

except Exception as e:
    print(f'❌ Falha ao iniciar cliente Oracle: {e}')

usuario = 'MAO8CT'
senha = '49l1)f=f3q6A'
string_conexao_completa = '(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=SI0EXARAC04.de.bosch.com)(PORT=38000))(CONNECT_DATA=(SERVICE_NAME=RLDP01_CON_4.BOSCH.COM)))'

try:
    connection = oracledb.connect(
        user=usuario,
        password=senha,
        dsn=string_conexao_completa
    )

    print('🔥 SUCESSO! Conectado com criptografia ativa ao banco da empresa!')
    
    cursor = connection.cursor()
    cursor.execute('SELECT user FROM dual')
    print(f'Resultado do teste (Usuário Ativo): {cursor.fetchone()[0]}')
    
    cursor.close()
    connection.close()
    print('Conexão encerrada com segurança.')

except Exception as e:
    print(f'❌ Erro na conexão: {e}')