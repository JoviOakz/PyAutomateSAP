import oracledb
import os

os.environ['TNS_ADMIN'] = r'C:\Oracle\ora_net'

alias = "REDLake_ZeusP_Consumer_Common.world"

try:
    connection = oracledb.connect(
        user="MAO8CT", 
        password="49l1)f=f3q6A", 
        dsn=alias
    )
    print("Conexão realizada com sucesso!")
    
    cursor = connection.cursor()
    cursor.execute("SELECT user FROM dual")
    print(f"Conectado como: {cursor.fetchone()[0]}")
    
    cursor.close()
    connection.close()

except Exception as e:
    print(f"Erro ao conectar: {e}")