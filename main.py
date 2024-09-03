import os
import csv
import fitz  # PyMuPDF
import docx
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import shutil
from typing import List, Dict

PALAVRAS_CHAVE_FABRICANTES = {
    "Google": ["google", "google cloud certified"],
    "IBM": ["ibm", "ibm certified"],
    "AWS": ["aws", "aws certification"],
    "Red Hat": ["red hat", "red hat certified"],
    "Liferay": ["liferay", "liferay certified"],
    "Delphix": ["delphix"],
    "Oracle": ["oracle", "oracle certified"],
}
PALAVRAS_CHAVE_EXCLUSAO = ["alura", "udemy", "coursera", "curriculo", "currículo"]

def extrair_texto_de_pdf(caminho_arquivo: str, usar_ocr: bool = False) -> str:
    try:
        doc = fitz.open(caminho_arquivo)
        texto = "".join(pagina.get_text() for pagina in doc)
        return texto if texto or not usar_ocr else extrair_texto_de_pdf_com_ocr(caminho_arquivo)
    except Exception as e:
        print(f"Erro ao extrair texto de {caminho_arquivo}: {e}")
        return ""

def extrair_texto_de_pdf_com_ocr(caminho_arquivo: str) -> str:
    try:
        imagens = convert_from_path(caminho_arquivo)
        return "".join(pytesseract.image_to_string(imagem) for imagem in imagens)
    except Exception as e:
        print(f"Erro ao extrair texto de {caminho_arquivo} usando OCR: {e}")
        return ""

def extrair_texto_de_docx(caminho_arquivo: str) -> str:
    try:
        doc = docx.Document(caminho_arquivo)
        return "\n".join(paragrafo.text for paragrafo in doc.paragraphs)
    except Exception as e:
        print(f"Erro ao extrair texto de {caminho_arquivo}: {e}")
        return ""

def extrair_texto_de_imagem(caminho_arquivo: str) -> str:
    try:
        imagem = Image.open(caminho_arquivo)
        return pytesseract.image_to_string(imagem)
    except Exception as e:
        print(f"Erro ao extrair texto de {caminho_arquivo} usando OCR: {e}")
        return ""

def copiar_arquivo_para_pasta(caminho_arquivo: str, nome_funcionario: str, nome_pasta: str):
    novo_diretorio = os.path.join("certificados_por_fabricante", nome_funcionario, nome_pasta)
    os.makedirs(novo_diretorio, exist_ok=True)
    shutil.copy2(caminho_arquivo, os.path.join(novo_diretorio, os.path.basename(caminho_arquivo)))

def texto_contem_palavras_chave_exclusao(texto: str) -> bool:
    return any(palavra_chave in texto.lower() for palavra_chave in PALAVRAS_CHAVE_EXCLUSAO)

def determinar_fabricante(texto: str) -> str:
    for fabricante, palavras_chave in PALAVRAS_CHAVE_FABRICANTES.items():
        if any(palavra_chave in texto.lower() for palavra_chave in palavras_chave):
            return fabricante
    return None

def processar_arquivo(caminho_arquivo: str, nome_funcionario: str) -> Dict:
    extensao_arquivo = os.path.splitext(caminho_arquivo)[1].lower()

    extratores = {
        ".pdf": extrair_texto_de_pdf,
        ".docx": extrair_texto_de_docx,
        ".jpg": extrair_texto_de_imagem,
        ".jpeg": extrair_texto_de_imagem,
        ".png": extrair_texto_de_imagem
    }

    extrator = extratores.get(extensao_arquivo)
    if not extrator:
        return {"Tipo": "Não Suportado", "Funcionario": nome_funcionario, "Arquivo": os.path.basename(caminho_arquivo), "Caminho": caminho_arquivo}
    
    texto = extrator(caminho_arquivo)
    if texto_contem_palavras_chave_exclusao(texto):
        return {"Tipo": "Exclusão", "Funcionario": nome_funcionario, "Arquivo": os.path.basename(caminho_arquivo), "Caminho": caminho_arquivo}
    
    fabricante = determinar_fabricante(texto)
    if fabricante is None:
        return {"Tipo": "Não Classificado", "Funcionario": nome_funcionario, "Arquivo": os.path.basename(caminho_arquivo), "Caminho": caminho_arquivo}
    
    copiar_arquivo_para_pasta(caminho_arquivo, nome_funcionario, fabricante)
    return {
        "Tipo": "Classificado",
        "Funcionario": nome_funcionario,
        "Arquivo": os.path.basename(caminho_arquivo),
        "Caminho": caminho_arquivo,
        "Fabricante": fabricante,
        "Texto": texto
    }

def processar_arquivos(diretorio: str) -> Dict:
    dados_funcionarios = []
    total_arquivos = 0
    total_classificados = 0
    total_nao_classificados = 0
    total_exclusao = 0
    total_nao_suportados = 0
    contagem_fabricantes = {fabricante: 0 for fabricante in PALAVRAS_CHAVE_FABRICANTES}

    for root, dirs, files in os.walk(diretorio):
        nome_funcionario = os.path.basename(root)
        for file in files:
            caminho_arquivo = os.path.join(root, file)
            total_arquivos += 1
            if total_arquivos % 100 == 0:
                print(f"Arquivos processados: {total_arquivos}")

            resultado_arquivo = processar_arquivo(caminho_arquivo, nome_funcionario)
            if resultado_arquivo:
                tipo = resultado_arquivo.get("Tipo")
                if tipo == "Classificado":
                    total_classificados += 1
                    fabricante = resultado_arquivo.get("Fabricante")
                    if fabricante:
                        contagem_fabricantes[fabricante] += 1
                elif tipo == "Exclusão":
                    total_exclusao += 1
                elif tipo == "Não Classificado":
                    total_nao_classificados += 1
                elif tipo == "Não Suportado":
                    total_nao_suportados += 1

                dados_funcionarios.append(resultado_arquivo)

    print(f"\nTotal de arquivos encontrados: {total_arquivos}")
    print(f"Total de arquivos classificados: {total_classificados}")
    print(f"Total de arquivos sem palavras chaves: {total_nao_classificados}")
    print(f"Total de arquivos com palavras de exclusão: {total_exclusao}")
    print(f"Total de arquivos com formatos não suportados: {total_nao_suportados}")
    print(f"\nArquivos organizados por fabricante:")
    for fabricante, contagem in contagem_fabricantes.items():
        print(f"  {fabricante}: {contagem}")

    return dados_funcionarios

def salvar_em_csv(dados: List[Dict], arquivo_saida: str):
    with open(arquivo_saida, mode='w', newline='', encoding='utf-8') as arquivo_csv:
        nomes_colunas = ["Tipo", "Funcionario", "Arquivo", "Caminho", "Fabricante", "Texto"]
        escritor = csv.DictWriter(arquivo_csv, fieldnames=nomes_colunas)
        escritor.writeheader()
        escritor.writerows(dados)

def main():
    diretorio = './colaboradores-2'
    arquivo_saida = 'certificados_encontrados.csv'
    
    print(f"Iniciando o processamento da pasta '{diretorio}'...")
    dados_funcionarios = processar_arquivos(diretorio)
    salvar_em_csv(dados_funcionarios, arquivo_saida)
    print(f"Processo concluído! Resultados salvos em {arquivo_saida}")

if __name__ == "__main__":
    main()
