from pdfminer.high_level import extract_text
import os

def extrair_texto_pdfminer(caminho_do_pdf):
    try:
        texto_completo = extract_text(caminho_do_pdf)
        paginas = texto_completo.split('\f')  # Divide o texto completo em páginas
        if len(paginas) > 1:  # Verifica se há pelo menos duas páginas
            texto_pagina = paginas[1]  # Pega a segunda página (índice 1)
            texto_pagina = texto_pagina.replace('\n', ' ')
            return texto_pagina.strip()  # Remove espaços extras no início e no fim
        else:
            print(f"O PDF {caminho_do_pdf} não tem pelo menos duas páginas.")
            return ''
    except Exception as e:
        print(f"Erro ao abrir ou ler o PDF {caminho_do_pdf}: {e}")
        return ''

# Exemplo de uso
caminho_do_pdf = 'caminho/para/seu/arquivo.pdf'
texto_extraido = extrair_texto_pdfminer(caminho_do_pdf)
print(f"Texto extraído do PDF {caminho_do_pdf}: {texto_extraido[:500]}...")  # Mostra os primeiros 500 caracteres do texto extraído