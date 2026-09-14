import oracledb
import pandas as pd

CAMINHO_INSTANT_CLIENT = r'C:\oracle\instantclient_23_0'

try:
    oracledb.init_oracle_client(lib_dir=CAMINHO_INSTANT_CLIENT)
    
except Exception as e:
    print(f'Aviso/Erro ao inicializar o cliente Oracle: {e}')

USUARIO = 'MAO8CT'
SENHA = '49l1)f=f3q6A'
dsn = 'REDLake_ZeusP_Consumer_Common.world'

df_excel = pd.read_excel('PyOdbc.xlsx')
lp_list = df_excel['Projeto'].dropna().unique().tolist()

if not lp_list:
    print('Nenhum projeto encontrado na planilha.')
    exit()

lps_text = ", ".join([f"'{lp}'" for lp in lp_list])

try:
    with oracledb.connect(user=USUARIO, password=SENHA, dsn=dsn) as conn:
        with conn.cursor() as cursor:
            query = f'''
                SELECT
                    PROJ.RB04_YT3_QMNUM AS RS,
                    PROJ.PSPID_EDIT     AS PROJETO,
                    AFPO.AUFNR          AS DIAGRAMA_REDE,
                    EBAN.BANFN          AS RC,
                    EKKN.EBELN          AS PO,
                    EKKN.EBELP          AS ITEM
                FROM MARD_MDNA.V_CUSN_PROJ_B2 PROJ
                JOIN MARD_MDNA.V_CUSN_AFPO_B2 AFPO
                    ON  AFPO.MANDT = PROJ.MANDT
                    AND AFPO.AUFNR = REPLACE(PROJ.PSPID_EDIT, '-', '')
                JOIN MARD_MDNA.V_CUSN_EKKN_B2 EKKN
                    ON  EKKN.MANDT = AFPO.MANDT
                    AND EKKN.NPLNR = AFPO.AUFNR
                LEFT JOIN MARD_MDNA.V_CUSN_EBAN_B2 EBAN
                    ON  EBAN.MANDT = EKKN.MANDT
                    AND EBAN.EBELN = EKKN.EBELN
                    AND EBAN.EBELP = EKKN.EBELP
                WHERE PROJ.VBUKR = '9084'
                AND PROJ.WERKS = '6854'
                AND PROJ.PSPID_EDIT IN ({lps_text})
            '''

            cursor.execute(query)

            dados = cursor.fetchall()

            for linha in dados:
                print(linha)

except oracledb.Error as e:
    print(f'Erro ao conectar ou consultar a view do SAP: {e}')