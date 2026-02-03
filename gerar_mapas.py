import os
import xml.etree.ElementTree as ET
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from collections import defaultdict
from datetime import datetime
from reportlab.pdfbase.pdfmetrics import stringWidth

import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

XML_DIR = os.path.join(BASE_DIR, "xml_entrada")
PDF_DIR = os.path.join(BASE_DIR, "mapas_pdf")
LOGO_PATH = os.path.join(BASE_DIR, "logo", "friovel.png")
REGRAS_PATH = os.path.join(BASE_DIR, "regras", "clientes_individuais.xlsx")
CONVERSAO_PATH = os.path.join(BASE_DIR, "regras", "conversao_produtos.xlsx")

os.makedirs(PDF_DIR, exist_ok=True)

motoristas_individuais = {
    "PATO BRANCO": "Iloi",
    "FRANCISCO BELTRAO": "Josué"
}

motoristas_consolidados = {
    "PATO BRANCO": "Rogerio",
    "FRANCISCO BELTRAO": "José Dirceu"
}

regras_df = pd.read_excel(REGRAS_PATH)
regras_df["Cidade"] = regras_df["Cidade"].str.upper().str.strip()
regras_df["CNPJ"] = regras_df["CNPJ"].astype(str).str.replace(r"\D", "", regex=True)

if os.path.exists(CONVERSAO_PATH):
    conv_df = pd.read_excel(CONVERSAO_PATH)
    conv_df["Codigo"] = conv_df["Codigo"].astype(str).str.strip()
    conversao = {
        row["Codigo"]: (row["Tipo"], int(row["Un_por_embalagem"]))
        for _, row in conv_df.iterrows()
    }

else:
    conversao = {}

ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

mapas_individuais = defaultdict(lambda: defaultdict(lambda: {
    "produtos": defaultdict(lambda: {"qtd": 0, "codigo": ""}),
    "endereco": "",
    "cnpj": "",
    "pesoL": 0.0,
    "pesoB": 0.0
}))

mapas_consolidados = defaultdict(lambda: {
    "produtos": defaultdict(lambda: {"qtd": 0, "codigo": ""}),
    "pesoL": 0.0,
    "pesoB": 0.0
})

for arquivo in os.listdir(XML_DIR):
    if not arquivo.lower().endswith(".xml"):
        continue

    tree = ET.parse(os.path.join(XML_DIR, arquivo))
    root = tree.getroot()

    cidade = root.find(".//nfe:enderDest/nfe:xMun", ns).text.upper().strip()

    dest = root.find(".//nfe:dest", ns)
    cliente = dest.find("nfe:xNome", ns).text.upper().strip()

    cnpj_elem = dest.find("nfe:CNPJ", ns)
    cpf_elem = dest.find("nfe:CPF", ns)
    doc = cnpj_elem.text if cnpj_elem is not None else cpf_elem.text
    doc = "".join(filter(str.isdigit, doc))

    ender = dest.find("nfe:enderDest", ns)

    rua = ender.find("nfe:xLgr", ns).text if ender.find("nfe:xLgr", ns) is not None else ""
    numero = ender.find("nfe:nro", ns).text if ender.find("nfe:nro", ns) is not None else ""
    bairro = ender.find("nfe:xBairro", ns).text if ender.find("nfe:xBairro", ns) is not None else ""
    cep = ender.find("nfe:CEP", ns).text if ender.find("nfe:CEP", ns) is not None else ""
    uf = ender.find("nfe:UF", ns).text if ender.find("nfe:UF", ns) is not None else ""

    endereco = f"{rua}, {numero} - {bairro} - {cidade}/{uf} - CEP {cep}"

    vol = root.find(".//nfe:transp/nfe:vol", ns)
    pesoL = float(vol.find("nfe:pesoL", ns).text) if vol is not None and vol.find("nfe:pesoL", ns) is not None else 0.0
    pesoB = float(vol.find("nfe:pesoB", ns).text) if vol is not None and vol.find("nfe:pesoB", ns) is not None else 0.0

    individual = not regras_df[
        (regras_df["Cidade"] == cidade) &
        (regras_df["CNPJ"] == doc)
    ].empty

    for det in root.findall(".//nfe:det", ns):
        prod = det.find("nfe:prod", ns)
        codigo = prod.find("nfe:cProd", ns).text.strip()
        nome = prod.find("nfe:xProd", ns).text.upper().strip()
        qtd = float(prod.find("nfe:qCom", ns).text)

        if individual:
            mapas_individuais[cidade][cliente]["endereco"] = endereco
            mapas_individuais[cidade][cliente]["cnpj"] = doc
            item = mapas_individuais[cidade][cliente]["produtos"][nome]
            item["qtd"] += qtd
            item["codigo"] = codigo
            mapas_individuais[cidade][cliente]["pesoL"] += pesoL
            mapas_individuais[cidade][cliente]["pesoB"] += pesoB
        else:
            item = mapas_consolidados[cidade]["produtos"][nome]
            item["qtd"] += qtd
            item["codigo"] = codigo
            mapas_consolidados[cidade]["pesoL"] += pesoL
            mapas_consolidados[cidade]["pesoB"] += pesoB

def quebrar_texto(texto, largura, fonte, tamanho):
    palavras = texto.split()
    linhas = []
    linha = ""
    for p in palavras:
        teste = linha + (" " if linha else "") + p
        if stringWidth(teste, fonte, tamanho) <= largura:
            linha = teste
        else:
            linhas.append(linha)
            linha = p
    if linha:
        linhas.append(linha)
    return linhas

def formatar_quantidade(codigo, total_un):
    if codigo not in conversao:
        return f"{int(total_un)} UN", 0, int(total_un)

    tipo, fator = conversao[codigo]
    cx = int(total_un // fator)
    un = int(total_un % fator)

    if cx > 0 and un > 0:
        return f"{cx} {tipo} + {un} UN", cx, un
    if cx > 0:
        return f"{cx} {tipo}", cx, 0
    return f"{un} UN", 0, un

def rodape(c, pagina):
    c.setFont("Helvetica", 8)
    c.drawCentredString(A4[0] / 2, 1.2 * cm, f"Página {pagina}")

def gerar_pdf(cidade, motorista, cliente, cnpj, endereco, dados, caminho):
    produtos = dados["produtos"]
    pesoL = dados["pesoL"]
    pesoB = dados["pesoB"]

    c = canvas.Canvas(caminho, pagesize=A4)
    largura, altura = A4
    pagina = 1

    total_cx = 0
    total_un = 0

    def cabecalho():
        c.drawImage(ImageReader(LOGO_PATH), 1.2*cm, altura-4*cm, width=6*cm, height=3*cm, preserveAspectRatio=True)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, altura-5*cm, "MAPA DE SEPARAÇÃO DE PRODUTOS")
        c.setFont("Helvetica", 10)
        c.drawString(2*cm, altura-6*cm, f"Cidade: {cidade}")
        c.drawString(2*cm, altura-6.7*cm, f"Motorista: {motorista}")

    def cabecalho_tabela(y):
        c.setFont("Helvetica-Bold", 10)
        c.drawString(2*cm, y, "Produto")
        c.drawString(15.6*cm, y, "✔")
        c.drawRightString(18.8*cm, y, "Quantidade")
        c.setFont("Helvetica", 10)
        return y - 0.8*cm

    cabecalho()
    y = altura - 7.8*cm

    if cliente:
        c.drawString(2*cm, y, f"Cliente: {cliente}")
        y -= 0.7*cm
        c.drawString(2*cm, y, f"CNPJ: {cnpj}")
        y -= 0.7*cm
        c.drawString(2*cm, y, f"Endereço: {endereco}")
        y -= 1.2*cm
    else:
        y -= 0.6*cm

    c.drawRightString(
        largura - 2*cm,
        altura - 2*cm,
        f"Data: {datetime.now().strftime('%d/%m/%Y')}"
    )


    y -= 1.2*cm

    y = cabecalho_tabela(y)

    produto_x = 2.2*cm
    produto_largura = 12.3*cm
    check_x = 15.6*cm
    qtd_x = 18.8*cm
    linha_base = 0.6*cm

    for produto, dados_prod in sorted(produtos.items()):
        codigo = dados_prod["codigo"]
        qtd = dados_prod["qtd"]

        texto_produto = f"{codigo} - {produto}"
        linhas = quebrar_texto(texto_produto, produto_largura, "Helvetica", 10)
        altura_bloco = max(len(linhas) * linha_base + 0.4*cm, 1.1*cm)

        if y - altura_bloco < 3*cm:
            rodape(c, pagina)
            c.showPage()
            pagina += 1
            cabecalho()
            y = altura - 7.8*cm
            y = cabecalho_tabela(y)

        c.rect(2*cm, y-altura_bloco+0.1*cm, 17*cm, altura_bloco)

        centro_y = y - altura_bloco/2
        texto_y = centro_y + (len(linhas)-1)*linha_base/2

        for linha in linhas:
            c.drawString(produto_x, texto_y, linha)
            texto_y -= linha_base

        c.rect(check_x, centro_y-0.25*cm, 0.45*cm, 0.45*cm)

        texto_qtd, cx, un = formatar_quantidade(codigo, qtd)
        total_cx += cx
        total_un += un

        c.drawRightString(qtd_x, centro_y-0.15*cm, texto_qtd)

        y -= altura_bloco

    base_y = max(y - 1.4*cm, 3*cm)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(2*cm, base_y, f"Totais: {total_cx} CX + {total_un} UN")

    c.setFont("Helvetica", 10)
    c.drawString(2*cm, base_y - 0.6*cm, f"Peso total (líquido): {pesoL:.2f} kg")
    c.drawString(2*cm, base_y - 1.2*cm, f"Peso bruto: {pesoB:.2f} kg")

    linha_y = base_y - 0.6*cm
    texto_assinatura_y = linha_y - 0.5*cm

    linha_inicio_x = 11.5*cm
    linha_fim_x = 18.8*cm

    c.line(linha_inicio_x, linha_y, linha_fim_x, linha_y)

    centro_linha_x = (linha_inicio_x + linha_fim_x) / 2

    c.drawCentredString(
        centro_linha_x,
        texto_assinatura_y,
        "Separador / Conferente"
    )

    rodape(c, pagina)
    c.save()

for cidade, clientes in mapas_individuais.items():
    motorista = motoristas_individuais.get(cidade, "")
    for cliente, dados in clientes.items():
        nome_pdf = f"{cidade}_{motorista}_{cliente}.pdf".replace(" ", "_")
        gerar_pdf(
            cidade,
            motorista,
            cliente,
            dados["cnpj"],
            dados["endereco"],
            dados,
            os.path.join(PDF_DIR, nome_pdf)
        )

for cidade, dados in mapas_consolidados.items():
    motorista = motoristas_consolidados.get(cidade, "")
    nome_pdf = f"{cidade}_{motorista}_CONSOLIDADO.pdf".replace(" ", "_")
    gerar_pdf(
        cidade,
        motorista,
        None,
        None,
        None,
        dados,
        os.path.join(PDF_DIR, nome_pdf)
    )
