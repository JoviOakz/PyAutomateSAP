import pdfplumber
import re

with pdfplumber.open(r'GDS - Orç.06572-26.pdf') as pdf:
    table = pdf.pages[0].extract_table()
    
    item_1 = table[1][0]
    item_2 = table[2][0]
    text_1 = table[1][1]
    text_2 = table[2][1]
    qty_1 = table[1][2]
    qty_2 = table[2][2]
    usi_time_1 = re.sub(r'[^0-9,]', '', table[1][3])
    usi_time_2 = table[2][3]
    total_value_1 = table[1][5]
    total_value_2 = table[2][5]
    om = re.sub(r'[^0-9]', '', table[5][1])

    print('\n')

    for linha in table:
        print(linha)

print('\n==========================================================\n')

print(f'''Fornecedor: GDS 
Orç: 06572-26 - ITEM 0{item_1}
LP/OM: {om} 
Descrição: {text_1} 
Qtd.: {qty_1} peças 
Prazo: 07/08/2026
Thais Fischer
            
Tempo de usinagem: {usi_time_1} horas p/pç''')

print('\n==========================================================\n')

print(f'''Fornecedor: GDS 
Orç: 06572-26 - ITEM 0{item_2}
LP/OM: {om} 
Descrição: {text_2} 
Qtd.: {qty_2} peças 
Prazo: 07/08/2026
Thais Fischer
            
Tempo de usinagem: {usi_time_2} horas p/pç\n''')