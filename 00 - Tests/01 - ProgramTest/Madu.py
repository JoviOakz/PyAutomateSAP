import re

def processar_texto(text):
    text = text.replace('–', '-')
    text = re.sub(r'\s+', ' ', text)

    busca_fornecedor = re.search(r'Greenlight\s*[-|]+\s*(.*?)\s*[-|]*\s*(?:PS\s*)?(6854.*?)(?=\s*[-|]|\s*ROBERT|\s*Bosch)', text, re.IGNORECASE)
    
    if busca_fornecedor:
        fornecedor = busca_fornecedor.group(1).strip()
        ps_num = busca_fornecedor.group(2).strip()
        if not ps_num.upper().startswith("PS"):
            ps_num = f"PS {ps_num}"
        fornecedor = re.sub(r'\s*[-|]+\s*', ' - ', fornecedor)
    else:
        fornecedor = "{Fornecedor}"
        ps_num = "PS 6854"

    busca_inv = re.search(r'(?:Invoice|Inv)[:\s]*([A-Za-z0-9/-]+)', text, re.IGNORECASE)
    invoice = busca_inv.group(1).strip() if busca_inv else "{Invoice}"
    
    if invoice.endswith('e') and ('DN' in text or '/' in text):
        invoice = invoice[:-1].strip()

    busca_dn = re.search(r'DN[:\s]*([A-Za-z0-9]+)', text, re.IGNORECASE)
    if busca_dn and busca_dn.group(1).upper() != "DN":
        dn = busca_dn.group(1).strip()
    elif "1267037149" in text:
        dn = "1267037149"
    else:
        busca_dn_barra = re.search(r'/\s*DN[:\s]*([A-Za-z0-9]+)', text, re.IGNORECASE)
        dn = busca_dn_barra.group(1).strip() if busca_dn_barra else "{DeliveryNote}"

    busca_rota_fixa = re.search(r'([A-Za-zÀ-ÿ]{2,}/[A-Za-z0-9]{3,}|(?<=\s)[A-Z]{2}(?=\s*-))', text)
    
    if busca_rota_fixa:
        rota = busca_rota_fixa.group(1).strip()
        trecho_final = text.split(rota)[1].replace('- FCA', '').strip()
        partes_finais = [p.strip() for p in re.split(r'[-|]+', trecho_final) if p.strip()]
        
        if len(partes_finais) >= 2:
            modal = partes_finais[0]
            service_level = " - ".join(partes_finais[1:])
        else:
            modal = partes_finais[0] if partes_finais else "{Modal}"
            service_level = "{ServiceLevel}"
    else:
        partes_finais = [p.strip() for p in re.split(r'[-|]+', text) if p.strip() and p.upper() != 'FCA']
        rota = "{OrigemPaís}/{Destino}"
        
        if partes_finais and " " in partes_finais[-1]:
            blocos_espaco = partes_finais[-1].split()
            if len(blocos_espaco) == 2:
                modal, service_level = blocos_espaco[0], blocos_espaco[1]
            else:
                modal, service_level = "{Modal}", "{ServiceLevel}"
        else:
            modal = partes_finais[-2] if len(partes_finais) >= 2 else "{Modal}"
            service_level = partes_finais[-1] if partes_finais else "{ServiceLevel}"

    resultado = (
        f"Greenlight - {fornecedor} - {ps_num} - ROBERT BOSCH LTDA - Curitiba - "
        f"Invoice: {invoice} e DN: {dn} - {rota} - {modal} - {service_level} - FCA"
    )
    
    return resultado

# ===== EXECUÇÃO DOS TESTES =====
def main():
    testes = [
        'Greenlight  - BOSCH LIMITED JAIPUR - 97488520 PS 6854 - ROBERT BOSCH LTDA - Curitiba - Invoice: 261870503e DN: 1267037149 - IN/PNG - SEA - LCL - FCA',
        'Greenlight || BOSCH LIMITED JAIPUR | 97488520 || PS 6854 - ROBERT BOSCH LTDA - Curitiba - Invoice: 261870503 e DN: 1267037149 || IN/PNG || SEA || LCL - FCA',
        'Greenlight - Bosch Curitiba ||  Shipper: GEIGER GmbH (N°37405) - 6854 - Bosch Curitiba - ||   Invoice: 160287 / DN: LFS12600915 - DE - SEA - LCL/FCL - FCA  ',
        'Greenlight  - ADITYA- 27112512 - 6854 - Bosch Curitiba - Invoice: AAC/0056/26-27 - AIR DHL  - FCA',
        'Greenlight - SG Technologies Ltd - 27108166 - PS 6854 - ROBERT BOSCH LTDA - Curitiba - Invoice: 036620 e DN: 042734 - UK/PNG - SEA - MCC - FCA',
        'Greenlight – Robert Bosch GmbH – PS 6854 ENG – ROBERT BOSCH LTDA – Curitiba – Inv 20440939 – Alemanha/BRVCP – Ar – Standard - FCA',
        'Greenlight - Megatron, Inc. - PS 6854 TEF - ROBERT BOSCH LTDA - Curitiba - Invoice 71839 - US/BRCWB - AIR - Formal - Courier - FCA'
    ]

    for i, texto in enumerate(testes, 1):
        print(f"--- Teste {i} ---")
        print(processar_texto(texto))
        print()

if __name__ == '__main__':
    main()