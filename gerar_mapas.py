import os
import xml.etree.ElementTree as ET
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from collections import defaultdict
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XML_DIR = os.path.join(BASE_DIR, "xml_entrada")
PDF_DIR = os.path.join(BASE_DIR, "mapas_pdf")
LOGO_PATH = os.path.join(BASE_DIR, "logo", "friovel.png")
REGRAS_PATH = os.path.join(BASE_DIR, "regras", "clientes_individuais.xlsx")

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

ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

mapas_individuais = defaultdict(lambda: defaultdict(lambda: {"produtos": defaultdict(int), "endereco": "", "cnpj": ""}))
mapas_consolidados = defaultdict(lambda: defaultdict(int))

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

    individual = not regras_df[
        (regras_df["Cidade"] == cidade) &
        (regras_df["CNPJ"] == doc)
    ].empty

    for det in root.findall(".//nfe:det", ns):
        prod = det.find("nfe:prod", ns)
        nome = prod.find("nfe:xProd", ns).text.strip()
        qtd = float(prod.find("nfe:qCom", ns).text)

        if individual:
            mapas_individuais[cidade][cliente]["endereco"] = endereco
            mapas_individuais[cidade][cliente]["cnpj"] = doc
            mapas_individuais[cidade][cliente]["produtos"][nome] += qtd
        else:
            mapas_consolidados[cidade][nome] += qtd

def rodape(c, pagina):
    c.setFont("Helvetica", 8)
    c.drawCentredString(A4[0]/2, 1.2*cm, f"Página {pagina}")

def gerar_pdf(cidade, motorista, cliente, cnpj, endereco, produtos, caminho):
    c = canvas.Canvas(caminho, pagesize=A4)
    largura, altura = A4
    pagina = 1

    def cabecalho():
        logo = ImageReader(LOGO_PATH)
        c.drawImage(logo, 1.2*cm, altura-4*cm, width=6*cm, height=3*cm, preserveAspectRatio=True)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, altura-5*cm, "MAPA DE SEPARAÇÃO DE PRODUTOS")
        c.setFont("Helvetica", 10)
        c.drawString(2*cm, altura-6*cm, f"Cidade: {cidade}")
        c.drawString(2*cm, altura-6.7*cm, f"Motorista: {motorista}")

    cabecalho()

    y = altura-7.8*cm

    if cliente:
        c.drawString(2*cm, y, f"Cliente: {cliente}")
        y -= 0.7*cm
        c.drawString(2*cm, y, f"CNPJ: {cnpj}")
        y -= 0.7*cm
        c.drawString(2*cm, y, f"Endereço: {endereco}")
        y -= 1.2*cm
    else:
        y -= 0.6*cm

    c.drawString(2*cm, y, f"Data: {datetime.now().strftime('%d/%m/%Y')}")
    y -= 1.2*cm

    c.setFont("Helvetica-Bold", 10)
    c.rect(2*cm, y-0.6*cm, 17*cm, 0.9*cm)
    c.drawString(2.2*cm, y-0.4*cm, "Produto")
    c.drawRightString(18.8*cm, y-0.4*cm, "Quantidade")
    y -= 1.2*cm

    c.setFont("Helvetica", 10)
    linha_altura = 0.9*cm

    for produto, qtd in sorted(produtos.items()):
        if y < 3*cm:
            rodape(c, pagina)
            c.showPage()
            pagina += 1
            cabecalho()
            y = altura-7.8*cm

        c.rect(2*cm, y-linha_altura+0.1*cm, 17*cm, linha_altura)
        c.drawString(2.2*cm, y-0.45*cm, produto[:90])
        c.drawRightString(18.8*cm, y-0.45*cm, f"{int(qtd)}")
        y -= linha_altura + 0.15*cm

    rodape(c, pagina)
    c.save()

for cidade, clientes in mapas_individuais.items():
    motorista = motoristas_individuais.get(cidade, "")
    for cliente, dados in clientes.items():
        produtos = dados["produtos"]
        endereco = dados["endereco"]
        cnpj = dados["cnpj"]
        nome_pdf = f"{cidade}_{motorista}_{cliente}.pdf".replace(" ", "_")
        gerar_pdf(cidade, motorista, cliente, cnpj, endereco, produtos, os.path.join(PDF_DIR, nome_pdf))

for cidade, produtos in mapas_consolidados.items():
    motorista = motoristas_consolidados.get(cidade, "")
    nome_pdf = f"{cidade}_{motorista}_CONSOLIDADO.pdf".replace(" ", "_")
    gerar_pdf(cidade, motorista, None, None, None, produtos, os.path.join(PDF_DIR, nome_pdf))
