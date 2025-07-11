from pyrfc import Connection
import json

# Configuração da conexão SAP
# Substitua pelos seus dados de conexão
config = {
    'user': 'ARA2CT',
    'passwd': 'AKDAgmr007@@y',
    'ashost': 'rb3ps0a0.server.bosch.com',
    'sysnr': '00',
    'client': '011',
    'lang': 'PT'
}

def conectar_sap():
    """Estabelece conexão com o SAP"""
    try:
        conn = Connection(**config)
        print("✓ Conexão estabelecida com sucesso!")
        return conn
    except Exception as e:
        print(f"✗ Erro ao conectar: {e}")
        return None

def buscar_ordem_completa(conn, numero_ordem):
    """Busca todos os dados da ordem especificada"""
    try:
        print(f"\n📋 Buscando dados da ordem: {numero_ordem}")
        
        # Chama a BAPI para buscar todos os detalhes da ordem
        ORDER_DETAIL = conn.call("BAPI_ALM_ORDER_GET_DETAIL", 
                                NUMBER=numero_ordem)
        
        return ORDER_DETAIL
    
    except Exception as e:
        print(f"✗ Erro ao buscar ordem: {e}")
        return None

def exibir_dados_ordem(order_data):
    """Exibe os dados da ordem de forma organizada"""
    if not order_data:
        print("Nenhum dado encontrado.")
        return
    
    print("\n" + "="*50)
    print("DADOS DA ORDEM - EQUIVALENTE IW23")
    print("="*50)
    
    # Cabeçalho
    if 'ES_HEADER' in order_data:
        header = order_data['ES_HEADER']
        print("\n📊 CABEÇALHO DA ORDEM:")
        print(f"  Número da Ordem: {header.get('AUFNR', 'N/A')}")
        print(f"  Tipo de Ordem: {header.get('AUART', 'N/A')}")
        print(f"  Status: {header.get('STAT', 'N/A')}")
        print(f"  Centro de Planejamento: {header.get('IWERK', 'N/A')}")
        print(f"  Descrição: {header.get('KTEXT', 'N/A')}")
        print(f"  Data de Criação: {header.get('ERDAT', 'N/A')}")
        print(f"  Criado por: {header.get('ERNAM', 'N/A')}")
        print(f"  Prioridade: {header.get('PRIOK', 'N/A')}")
    
    # Operações
    if 'ET_OPERATIONS' in order_data and order_data['ET_OPERATIONS']:
        print("\n🔧 OPERAÇÕES:")
        for i, op in enumerate(order_data['ET_OPERATIONS']):
            print(f"  Operação {i+1}:")
            print(f"    Número: {op.get('VORNR', 'N/A')}")
            print(f"    Descrição: {op.get('LTXA1', 'N/A')}")
            print(f"    Centro de Trabalho: {op.get('ARBPL', 'N/A')}")
            print(f"    Duração: {op.get('DAUNO', 'N/A')} {op.get('DAUNE', 'N/A')}")
    
    # Componentes
    if 'ET_COMPONENTS' in order_data and order_data['ET_COMPONENTS']:
        print("\n📦 COMPONENTES/MATERIAIS:")
        for i, comp in enumerate(order_data['ET_COMPONENTS']):
            print(f"  Componente {i+1}:")
            print(f"    Material: {comp.get('MATNR', 'N/A')}")
            print(f"    Descrição: {comp.get('MAKTX', 'N/A')}")
            print(f"    Quantidade: {comp.get('BDMNG', 'N/A')} {comp.get('MEINS', 'N/A')}")
            print(f"    Depósito: {comp.get('LGORT', 'N/A')}")
    
    # Objetos Técnicos
    if 'ET_OBJECTS' in order_data and order_data['ET_OBJECTS']:
        print("\n🏭 OBJETOS TÉCNICOS:")
        for i, obj in enumerate(order_data['ET_OBJECTS']):
            print(f"  Objeto {i+1}:")
            print(f"    Equipamento: {obj.get('EQUNR', 'N/A')}")
            print(f"    Local de Instalação: {obj.get('TPLNR', 'N/A')}")
            print(f"    Descrição: {obj.get('PLTXT', 'N/A')}")
    
    # Parceiros
    if 'ET_PARTNERS' in order_data and order_data['ET_PARTNERS']:
        print("\n👥 PARCEIROS:")
        for i, partner in enumerate(order_data['ET_PARTNERS']):
            print(f"  Parceiro {i+1}:")
            print(f"    Função: {partner.get('PARVW', 'N/A')}")
            print(f"    Número: {partner.get('PARNR', 'N/A')}")
            print(f"    Nome: {partner.get('NAME1', 'N/A')}")
    
    # Textos
    if 'ET_TEXTS' in order_data and order_data['ET_TEXTS']:
        print("\n📝 TEXTOS/OBSERVAÇÕES:")
        for i, text in enumerate(order_data['ET_TEXTS']):
            print(f"  Texto {i+1}:")
            print(f"    Tipo: {text.get('TDID', 'N/A')}")
            print(f"    Conteúdo: {text.get('LINE', 'N/A')}")
    
    # Confirmações
    if 'ET_CONFIRMATIONS' in order_data and order_data['ET_CONFIRMATIONS']:
        print("\n✅ CONFIRMAÇÕES:")
        for i, conf in enumerate(order_data['ET_CONFIRMATIONS']):
            print(f"  Confirmação {i+1}:")
            print(f"    Data: {conf.get('BUDAT', 'N/A')}")
            print(f"    Horas: {conf.get('ISMNW', 'N/A')}")
            print(f"    Usuário: {conf.get('ERNAM', 'N/A')}")

def salvar_dados_json(order_data, numero_ordem):
    """Salva os dados em arquivo JSON para análise posterior"""
    try:
        filename = f"ordem_{numero_ordem}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(order_data, f, indent=2, ensure_ascii=False, default=str)
        print(f"\n💾 Dados salvos em: {filename}")
    except Exception as e:
        print(f"✗ Erro ao salvar arquivo: {e}")

def main():
    # Número da ordem que você quer acessar
    NUMERO_ORDEM = "14081538"
    
    # Conecta ao SAP
    conn = conectar_sap()
    if not conn:
        return
    
    try:
        # Busca dados da ordem
        order_data = buscar_ordem_completa(conn, NUMERO_ORDEM)
        
        if order_data:
            # Exibe os dados
            exibir_dados_ordem(order_data)
            
            # Salva em arquivo JSON
            salvar_dados_json(order_data, NUMERO_ORDEM)
            
            print(f"\n✓ Processamento concluído para ordem {NUMERO_ORDEM}")
        else:
            print(f"✗ Ordem {NUMERO_ORDEM} não encontrada ou sem dados")
    
    finally:
        # Fecha a conexão
        conn.close()
        print("\n🔌 Conexão fechada")

if __name__ == "__main__":
    main()