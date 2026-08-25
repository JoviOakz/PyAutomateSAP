import oracledb

CAMINHO_INSTANT_CLIENT = r"C:\oracle\instantclient_23_0"

try:
    oracledb.init_oracle_client(lib_dir=CAMINHO_INSTANT_CLIENT)
    
except Exception as e:
    print(f"Aviso/Erro ao inicializar o cliente Oracle: {e}")

USUARIO = "MAO8CT"
SENHA = "49l1)f=f3q6A"
dsn = "REDLake_ZeusP_Consumer_Common.world"

try:
    with oracledb.connect(user=USUARIO, password=SENHA, dsn=dsn) as conexao:
        with conexao.cursor() as cursor:
            query = """
                SELECT
                    PROJ.RB04_YT3_QMNUM AS NOTA_PM,
                    PROJ.PSPID_EDIT AS PROJETO,
                    EBAN.BANFN AS NUM_REQUISICAO,
                    EBAN.EBELN AS NUM_PEDIDO
                FROM MARD_MDNA.V_CUSN_PROJ_B2 PROJ
                LEFT JOIN MARD_MDNA.V_CUSN_EBAN_B2 EBAN
                    ON  EBAN.MANDT = PROJ.MANDT
                    AND EBAN.DISUB_PSPNR = PROJ.PSPNR
                WHERE PROJ.VBUKR = '9084'
                AND PROJ.WERKS = '6854'
                AND PROJ.RB04_YT3_QMNUM LIKE '%14227041%'
            """

            cursor.execute(query)

            dados = cursor.fetchall()

            for linha in dados[:20]:
                print(linha)

except oracledb.Error as e:
    print(f"Erro ao conectar ou consultar a view do SAP: {e}")