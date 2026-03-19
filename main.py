import pdfplumber
from docx import Document
import re

# =========================
# FUNÇÃO PARA EXTRAIR DADOS
# =========================
def extrair_dados(pdf_path):
    texto_completo = ""

    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo += texto + "\n"

    # DEBUG (se quiser ver o texto extraído)
    # print(texto_completo)

    dados = {}

    padroes = {
        "CNPJ": r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})",
        "Razão Social": r"EMPRESA\s*(.*?)\n",
        "Nome Fantasia": r"NOME FANTASIA\s*(.*?)\n",
        "Situação Cadastral": r"SITUAÇÃO CADASTRAL\s*(.*?)\n",
        "Data da Situação": r"DATA DA SITUAÇÃO CADASTRAL\s*(.*?)\n",
        "Logradouro": r"LOGRADOURO\s*(.*?)\n",
        "Número": r"NÚMERO\s*(.*?)\n",
        "Município": r"MUNICÍPIO\s*(.*?)\n",
        "UF": r"UF\s*(.*?)\n",
        "CNAE Principal": r"CÓDIGO E DESCRIÇÃO DA ATIVIDADE ECONÔMICA PRINCIPAL\s*(.*?)\n"
    }

    for campo, padrao in padroes.items():
        resultado = re.search(padrao, texto_completo, re.IGNORECASE)

        if resultado:
            try:
                dados[campo] = resultado.group(1).strip()
            except IndexError:
                dados[campo] = resultado.group(0).strip()
        else:
            dados[campo] = "Não encontrado"

    return dados


# =========================
# FUNÇÃO PARA GERAR WORD
# =========================
def gerar_word(dados, nome_arquivo="empresa.docx"):
    doc = Document()
    doc.add_heading('Dados da Empresa', 0)

    for chave, valor in dados.items():
        doc.add_paragraph(f"{chave}: {valor}", style='List Bullet')

    doc.save(nome_arquivo)
    print(f"\nDocumento gerado: {nome_arquivo}")


# =========================
# EXECUÇÃO
# =========================
if __name__ == "__main__":
    caminho_pdf = "cnpj.pdf"  # coloque o nome correto aqui

    dados_empresa = extrair_dados(caminho_pdf)

    print("\n📄 Dados extraídos:\n")
    for k, v in dados_empresa.items():
        print(f"{k}: {v}")

    gerar_word(dados_empresa)