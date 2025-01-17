# COORDENADAS PARA EXTRAÇÃO DE DADOS USANDO OCR

def corte_copergas():
    corte = {
    'cnpj': (2550, 1020, 3110, 1200),
    'cnpj_ajustado': (2478, 1011, 3110, 1165), #

    'valor_total': (1900, 1560, 2700, 1800), 
    'valor_total_ajustado':(),

    'volume_total': (2900, 1560, 3500, 1800),
    'volume_total_ajustado':  (2040, 2990, 2285, 3100),
    'volume_total_ajustado2': (2000, 2890, 2385, 3200), # 
    
    'data_emissao': (3100, 1195, 3800, 1300),
    'data_emissao2': (1200, 1240, 3800, 1450), 
    
    'data_inicio': (1000, 3450, 2060, 3520), #
    'data_inicio_ajustado': (1000, 3520, 2060, 3630),#
    'data_inicio_ajustado2': (1300, 3420, 2210, 3640), # 29/08
    'data_inicio_ajustado3': (1320, 3560, 1950, 3622),#
    'data_inicio_ajustado4': (1330, 3310, 2000, 3685),#  
    'data_inicio_ajustado5': (1000, 3380, 2170, 3665),# 

    'data_fim': (1000, 3520, 2060, 3630),
    'data_fim_ajustado': (252, 3565, 610, 3620), # 29/08
    'data_fim_ajustado2': (1300, 3675, 2010, 3740),# 29/08
    'data_fim_ajustado3': (2000, 3450, 2560, 3520),# 29/08
    'data_fim_ajustado4': (241, 3508, 595, 3564),# 29/08
    'data_fim_ajustado5': (500, 3455, 1200, 3670),# 29/08
    
    'numero_fatura': (1340, 590, 1900, 700),#
    'numero_fatura_ajustado': (000, 430, 3750, 590),
    'numero_fatura_ajustado2': (1355, 610, 1900, 700), #

    'valor_icms': (900, 1800, 1350, 1970),
    'valor_icms_ajustado': (890, 1800, 1400, 2000),
    'valor_icms_ajustado2': (850, 1790, 1420, 2020),#
    'valor_icms_ajustado3': (945, 1950, 1400, 2020),

    }
    return corte

caminho_excel = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\dados.xlsx'

'''while True
    for Player(get_moeda):
        if player(moeda >= 3):
            brilhar(dourado)
        else:
            brilhar(branco)'''


