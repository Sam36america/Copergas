import os
import PyPDF2

def extrair_texto_dos_pdfs(pasta):
    textos = {}
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            caminho_do_pdf = os.path.join(pasta, arquivo)
            texto = extrair_texto(caminho_do_pdf)
            textos[arquivo] = texto
    return textos

def extrair_texto(caminho_do_pdf):
    texto = ''
    try:
        with open(caminho_do_pdf, 'rb') as arquivo:
            leitor_pdf = PyPDF2.PdfReader(arquivo)
            for pagina in leitor_pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_pagina = texto_pagina.replace('\n', ' ')
                    texto += texto_pagina + ' '
    except Exception as e:
        print(f"Erro ao abrir ou ler o PDF {caminho_do_pdf}: {e}")
    
    if not texto:
        print(f"Erro ao extrair texto do PDF: {caminho_do_pdf}")
    else:
        print(f"Texto extraído do PDF {caminho_do_pdf}: {texto[:500]}...")  # Mostra os primeiros 500 caracteres do texto extraído
    return texto.strip()  # Remove espaços extras no início e no fim

# Exemplo de uso
pasta = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Naturgy RJ\Faturas'
textos_extraidos = extrair_texto_dos_pdfs(pasta)
for nome_arquivo, texto in textos_extraidos.items():
    print(f"Texto extraído do arquivo {nome_arquivo}: {texto[:500]}...")  # Mostra os primeiros 500 caracteres do texto extraído