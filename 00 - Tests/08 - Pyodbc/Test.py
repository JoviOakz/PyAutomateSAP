import pyodbc

# Configuração da conexão com o banco do SAP (Exemplo com SAP HANA)
dados_conexao = (
    "DRIVER={HDBODBC};"  # Driver do SAP HANA (ou ODBC Driver para SQL Server)
    "SERVERNODE=seu_servidor_sap:30015;"  # IP/Host e Porta do banco
    "SERVERDB=NOME_DO_BANCO;"
    "UID=seu_usuario_banco;"
    "PWD=sua_senha_banco;"
)

try:
    with pyodbc.connect(dados_conexao) as conexao:
        with conexao.cursor() as cursor:
            # Query buscando dados de uma transação contábil específica (Ex: FB03)
            # BKPF é a tabela SAP de Cabeçalho de Documento de Contabilidade
            query = """
                SELECT TOP 100 COMPANY_CODE, DOC_NUMBER, FISCAL_YEAR, DOC_DATE, USER_NAME
                FROM BKPF
                WHERE FISCAL_YEAR = ? AN\D COMPANY_CODE = ?
            """

            # Executa passando o Ano Fiscal (2026) e Empresa (1000) como filtros
            cursor.execute(query, ("2026", "6854"))

            # Coleta e exibe os dados
            transacoes = cursor.fetchall()
            print(
                f"{"Empresa":<8} | {"Nº Documento":<12} | {"Ano":<6} | {"Data":<10} | {"Usuário":<10}"
            )
            print("-" * 55)
            for linha in transacoes:
                print(
                    f"{linha[0]:<8} | {linha[1]:<12} | {linha[2]:<6} | {linha[3]} | {linha[4]:<10}"
                )

except pyodbc.Error as e:
    print(f"Erro ao conectar ou consultar o banco SAP: {e}")