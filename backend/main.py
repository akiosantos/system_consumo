from email.utils import parseaddr
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import imaplib
import email
import os
from email.header import decode_header
from pypdf import PdfReader, PdfWriter
from fastapi.responses import FileResponse, StreamingResponse
import pdfplumber
import shutil
import re
import csv
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

IMAP_SERVER = "mail.barueri.sp.gov.br"

SYSTEM_DIR = r"U:\BackupContabilidade\Custos\0 - Enel, Sabesp e Telefônica - Lucas\system"
BASE_DIR = os.path.join(SYSTEM_DIR, "backend")
FRONT_DIR = os.path.join(SYSTEM_DIR, "frontend")

PLANILHA_ENEL = r"U:\BackupContabilidade\Custos\0 - Enel, Sabesp e Telefônica - Lucas\2026\1 - ENEL\1 - ENEL 2026.xlsx"
ABA_ENEL = "ENEL 2026"

PLANILHA_SABESP = r"U:\BackupContabilidade\Custos\0 - Enel, Sabesp e Telefônica - Lucas\2026\3 - Sabesp\Sabesp 2026.xlsx"
ABA_SABESP = "SABESP 2026"

PDF_SABESP_COM_CODIGO = os.path.join(BASE_DIR, "sabesp_com_codigo.pdf")

# ===== SABESP =====
EMAIL_SABESP = os.getenv("EMAIL_SABESP", "e-mail")
SENHA_SABESP = os.getenv("SENHA_SABESP", "sua senha")
REMETENTES_SABESP = [
    "e-mail remetente",
    "e-mail remetente"
]

PASTA_SABESP = os.path.join(BASE_DIR, "sabesp_pdf")
PASTA_SABESP_SEM_SENHA = os.path.join(BASE_DIR, "sabesp_pdf_sem_senha")
CSV_SABESP = os.path.join(BASE_DIR, "sabesp_consolidado.csv")
PDF_SABESP_COMPLETO = os.path.join(BASE_DIR, "sabesp_completo.pdf")
SENHAS_SABESP = ["465", "MIG"]

# ===== ENEL =====
EMAIL_ENEL = os.getenv("EMAIL_ENEL", "e-mail")
SENHA_ENEL = os.getenv("SENHA_ENEL", "sua senha")
REMETENTE_ENEL = [
    "e-mail remetente",
    "e-mail remetente"
]

PASTA_ENEL = os.path.join(BASE_DIR, "enel_pdf")
PASTA_ENEL_SEM_SENHA = os.path.join(BASE_DIR, "enel_pdf_sem_senha")
PDF_ENEL_FILTRADO = os.path.join(BASE_DIR, "enel_filtrado.pdf")
PDF_ENEL_COM_CODIGO = os.path.join(BASE_DIR, "enel_com_codigo.pdf")
CSV_ENEL = os.path.join(BASE_DIR, "enel_consolidado.csv")
SENHA_ENEL_PDF = os.getenv("SENHA_ENEL_PDF", "46523")

os.makedirs(PASTA_SABESP, exist_ok=True)
os.makedirs(PASTA_SABESP_SEM_SENHA, exist_ok=True)
os.makedirs(PASTA_ENEL, exist_ok=True)
os.makedirs(PASTA_ENEL_SEM_SENHA, exist_ok=True)


# ================= UTIL =================
def decodificar(texto):
    partes = decode_header(texto)
    return "".join(
        parte.decode(enc or "utf-8", errors="ignore") if isinstance(parte, bytes) else parte
        for parte, enc in partes
    )

def normalizar(texto):
    return re.sub(r"\s+", " ", texto.lower()) if texto else ""

def normalizar_instalacao(valor):
    return valor.lstrip("0") if valor else ""


# ================= PDF =================
def tentar_remover_senha(caminho_entrada, caminho_saida, senhas):
    reader = PdfReader(caminho_entrada)

    if not reader.is_encrypted:
        shutil.copy2(caminho_entrada, caminho_saida)
        return True

    ok = False
    for senha in senhas:
        try:
            if reader.decrypt(senha) != 0:
                ok = True
                break
        except:
            continue

    if not ok:
        return False

    writer = PdfWriter()
    writer.append(reader)

    with open(caminho_saida, "wb") as f:
        writer.write(f)

    return True

def juntar_pdfs(lista_pdfs, pdf_saida):
    writer = PdfWriter()

    for caminho in lista_pdfs:
        if not os.path.exists(caminho):
            print(f"[JUNTAR] Arquivo nao encontrado: {caminho}")
            continue

        try:
            reader = PdfReader(caminho)
            for page in reader.pages:
                nova_pagina = writer.add_page(page)
                nova_pagina.mediabox = page.mediabox
                if "/CropBox" in page:
                    nova_pagina.cropbox = page.cropbox
        except Exception as e:
            print(f"[JUNTAR] Erro ao abrir {caminho}: {e}")

    if len(writer.pages) == 0:
        print("[JUNTAR] Nenhum PDF valido encontrado.")
        return False

    with open(pdf_saida, "wb") as f:
        writer.write(f)

    return True

# ================= DEDUPLICAÇÃO SABESP =================
def deduplificar_pdf_sabesp(pdf_entrada, pdf_saida):
    """
    Lê o PDF completo, agrupa páginas por fatura (bloco iniciado por 'FATURAMENTO'),
    e escreve apenas um bloco por combinação única de fornecimento+vencimento+valor.
    """
    reader = PdfReader(pdf_entrada)
    
    # 1. Agrupa páginas em blocos de fatura
    blocos = []       # cada item: {"chave": str, "paginas": [page, ...]}
    bloco_atual = None

    for page in reader.pages:
        texto = page.extract_text() or ""
        texto_upper = texto.upper()

        if "FATURAMENTO" in texto_upper:
            # Extrai dados de identificação da fatura
            fornecimento = extrair_fornecimento_sabesp(texto)
            
            venc = re.search(r'VENCIMENTO:\s*(\d{2}/\d{2}/\d{4})', texto)
            vencimento = venc.group(1) if venc else ""

            texto_upper2 = texto.upper()
            m_total = re.search(r"TOTAL[\s\S]{0,60}?R\$[\s\*]*([\d.,]+)", texto_upper2)
            if m_total:
                valor = m_total.group(1)
            else:
                texto_limpo = texto.replace("*", "")
                valores = re.findall(r"R\$\s*([\d.,]+)", texto_limpo)
                if valores:
                    valores_float = [float(v.replace(".", "").replace(",", ".")) for v in valores]
                    valor = f"{max(valores_float):.2f}".replace(".", ",")
                else:
                    valor = ""

            chave = f"{fornecimento}_{vencimento}_{valor}"
            bloco_atual = {"chave": chave, "paginas": [page]}
            blocos.append(bloco_atual)
        else:
            # Página de continuação (código de barras, etc.)
            if bloco_atual is not None:
                bloco_atual["paginas"].append(page)
            # Se não há bloco_atual (página solta no início), ignora

    # 2. Filtra blocos únicos, mantendo a primeira ocorrência
    vistos = set()
    writer = PdfWriter()
    duplicatas = 0

    for bloco in blocos:
        chave = bloco["chave"]
        if chave in vistos:
            print(f"[DEDUP] Duplicata removida: {chave}")
            duplicatas += 1
            continue
        vistos.add(chave)
        for page in bloco["paginas"]:
            writer.add_page(page)

    print(f"[DEDUP] Total de blocos: {len(blocos)} | Duplicatas removidas: {duplicatas} | Únicos: {len(vistos)}")

    if len(writer.pages) == 0:
        raise ValueError("Nenhuma fatura única encontrada após deduplicação.")

    with open(pdf_saida, "wb") as f:
        writer.write(f)

    return duplicatas


# ================= SABESP EXTRACAO =================
def extrair_consumo_sabesp(texto):
    texto = re.sub(r"\s+", " ", texto)

    m = re.search(r"\d{2}/\d{2}/\d{2}\s+\d{2}/\d{2}/\d{2}\s+(\d{1,6})\s+\d{1,6}", texto)
    if m:
        return m.group(1)

    m = re.search(r"\d{2}/\d{2}/\d{2}\s+\d{1,6}\s+(\d{1,6})\s+\d{1,6}", texto)
    if m:
        return m.group(1)

    m = re.search(r"\d{2}/\d{2}/\d{2}.*?\d{1,6}\s+(\d{1,6})\s+\d{1,6}", texto)
    if m:
        return m.group(1)

    return ""

def extrair_fornecimento_sabesp(texto):
    m = re.search(r"\b(\d{9,16})\b", texto)
    if m:
        return m.group(1)
    return ""

def extrair_dados_sabesp(pdf):
    registros = []
    faturas_processadas = set()

    reader = PdfReader(pdf)
    contador_sequencial = 1

    for i, page in enumerate(reader.pages):
        texto = page.extract_text() or ""

        if "FATURAMENTO" not in texto.upper():
            continue

        fornecimento = extrair_fornecimento_sabesp(texto)

        vencimento_match = re.search(r'VENCIMENTO:\s*(\d{2}/\d{2}/\d{4})', texto)
        vencimento = vencimento_match.group(1) if vencimento_match else ""

        if not fornecimento:
            continue

        texto_upper = texto.upper()
        m_total = re.search(r"TOTAL[\s\S]{0,60}?R\$[\s\*]*([\d.,]+)", texto_upper)

        if m_total:
            valor = m_total.group(1)
        else:
            texto_limpo = texto.replace("*", "")
            valores = re.findall(r"R\$\s*([\d.,]+)", texto_limpo)
            if valores:
                valores_float = [float(v.replace(".", "").replace(",", ".")) for v in valores]
                maior = max(valores_float)
                valor = f"{maior:.2f}".replace(".", ",")
            else:
                valor = ""

        retencao = re.search(r'Reten.{1,3}o:\s*4,8%\s*([\d.,]+)', texto)
        retencao = retencao.group(1) if retencao else ""

        consumo = extrair_consumo_sabesp(texto)

        id_fatura_completo = f"{fornecimento}_{vencimento}_{valor}_{consumo}"

        if id_fatura_completo in faturas_processadas:
            continue

        registros.append([
            f"Pagina {i+1}",
            fornecimento,
            vencimento,
            consumo,
            valor,
            retencao
        ])

        faturas_processadas.add(id_fatura_completo)
        contador_sequencial += 1

    with open(CSV_SABESP, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina", "Fornecimento", "Vencimento", "Consumo_M3", "Valor_Total", "Retencao_4_8"])
        writer.writerows(registros)


# ================= PLANILHA =================
def carregar_mapa_instalacao_codigo():
    wb = load_workbook(PLANILHA_ENEL, data_only=True)
    ws = wb[ABA_ENEL]

    mapa = {}
    for linha in ws.iter_rows(min_row=2):
        codigo = linha[1].value
        instalacao = linha[3].value
        if codigo and instalacao:
            mapa[normalizar_instalacao(str(instalacao))] = str(codigo)

    print("Mapa instalacao -> codigo:", mapa)
    return mapa

def carregar_mapa_fornecimento_codigo_sabesp():
    wb = load_workbook(PLANILHA_SABESP, data_only=True)
    ws = wb[ABA_SABESP]

    mapa = {}
    for linha in ws.iter_rows(min_row=2):
        codigo = linha[1].value
        fornecimento = linha[4].value
        if codigo and fornecimento:
            fornecimento = str(fornecimento).strip()
            mapa[fornecimento] = str(codigo)

    print("Mapa SABESP fornecimento -> codigo:", mapa)
    return mapa


# ================= ENEL =================
def pagina_eh_fatura(texto):
    t = texto.lower()

    palavras_excluir = [
        "faturamento a menor",
        "aus",
        "assunto:",
        "parcelado em",
        "diferenca em relacao",
        "consumo acumulado",
        "aqui, voce pode acompanhar",
    ]

    for p in palavras_excluir:
        if p in t:
            return False

    pontos = 0
    if "instala" in t or "uc" in t:
        pontos += 1
    if "vencimento" in t:
        pontos += 1
    if re.search(r"r\$\s*\d", t):
        pontos += 1
    if "total" in t:
        pontos += 1

    return pontos >= 3


def filtrar_pdf_enel(pdf_entrada, pdf_saida):
    reader = PdfReader(pdf_entrada)
    writer = PdfWriter()

    for page in reader.pages:
        texto = page.extract_text() or ""
        if pagina_eh_fatura(texto):
            writer.add_page(page)

    if len(writer.pages) == 0:
        raise ValueError("Nenhuma pagina de fatura encontrada no PDF da Enel apos filtro.")

    with open(pdf_saida, "wb") as f:
        writer.write(f)

def extrair_instalacao(texto):
    m_mte = re.search(r"\b(MTE[A-Z0-9]{5,15})\b", texto, re.IGNORECASE)
    if m_mte:
        return m_mte.group(1).upper()

    m_num = re.search(r"\b\d{8,12}\b", texto)
    if m_num:
        return normalizar_instalacao(m_num.group(0))

    return ""

def extrair_referencia(texto, instalacao):
    if not instalacao:
        return ""

    t = re.sub(r"\s+", " ", texto)
    instalacao = instalacao.lstrip("0")

    pos = -1
    for padrao in [instalacao, instalacao.zfill(len(instalacao)+1)]:
        p = t.find(padrao)
        if p != -1:
            pos = p
            break

    area = t[pos:pos+500] if pos != -1 else t
    area = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", area)
    m = re.search(r"\b(0[1-9]|1[0-2])/[0-9]{4}\b", area)

    return m.group(0) if m else ""

def extrair_total(texto):
    texto = texto.lower()

    if re.search(r"r\$\s*\*+", texto):
        return "0,00"

    m = re.search(r"r\$\s*([\d.,]+)", texto)
    if m:
        return m.group(1)

    return ""

def extrair_ir(texto):
    texto = texto.lower()

    m = re.search(
        r"ret\.\s*art\.\s*64\s*lei\s*9430\s*-\s*1[,\.]20%\s*(?:[\d.,]+\s*){0,3}(-?\d[\d.,]*)",
        texto
    )
    if m:
        return m.group(1).replace("-", "")

    m = re.search(r"irrf?\s*1[,\.]20\s*%\s*r?\$?\s*(-?\d[\d.,]*)", texto)
    if m:
        return m.group(1).replace("-", "")

    return "0,00"

def extrair_consumo_enel(texto):
    texto = texto.upper()
    valores = []

    padrao_especial = re.findall(
        r"EN (CONSUMIDA|FORNECIDA)\s+(?:FAT\s+)?TU\s+KWH\s+([\d.,]+)",
        texto
    )
    if padrao_especial:
        for _, v in padrao_especial:
            numero = float(v.replace(".", "").replace(",", "."))
            valores.append(numero)

    if not valores:
        m = re.search(r"(?:CONSUMO|USO SIST\. DISTR\.) .*?KWH\s+([\d.,]+)", texto)
        if m:
            numero = float(m.group(1).replace(".", "").replace(",", "."))
            valores.append(numero)

    if not valores:
        return ""

    total = sum(valores)
    return f"{total:.2f}".replace(".", ",")

def escrever_codigo_e_ordenar(pdf_entrada, pdf_saida, mapa):
    reader = PdfReader(pdf_entrada)
    paginas = []

    for page in reader.pages:
        texto = page.extract_text() or ""
        inst = extrair_instalacao(normalizar(texto))
        codigo = mapa.get(inst)

        if codigo:
            packet = BytesIO()

            if "/CropBox" in page:
                del page["/CropBox"]

            if "/CropBox" in page:
                box = page["/CropBox"]
            else:
                box = page["/MediaBox"]

            largura = float(box[2]) - float(box[0])
            altura = float(box[3]) - float(box[1])

            can = canvas.Canvas(packet, pagesize=(largura, altura))
            can.setFont("Helvetica-Bold", 12)
            can.drawString(largura - 150, altura - 45, f"COD: {codigo}")
            can.save()

            packet.seek(0)
            overlay = PdfReader(packet)

            page.merge_transformed_page(
                overlay.pages[0],
                [1, 0, 0, 1, 0, 0],
                expand=False
            )

        paginas.append((codigo or "99999", page))

    paginas.sort(key=lambda x: [int(s) if s.isdigit() else s for s in re.findall(r"\d+|[A-Z]+", x[0])])

    writer = PdfWriter()
    for _, page in paginas:
        writer.add_page(page)

    with open(pdf_saida, "wb") as f:
        writer.write(f)

def escrever_codigo_e_ordenar_sabesp(pdf_entrada, pdf_saida, mapa):
    reader = PdfReader(pdf_entrada)
    paginas = []
    codigo_atual = None

    for page in reader.pages:
        texto = page.extract_text() or ""
        texto_upper = texto.upper()

        fornecimento = extrair_fornecimento_sabesp(texto)

        if "FATURAMENTO" in texto_upper:
            codigo_atual = None
            if fornecimento and fornecimento in mapa:
                codigo_atual = mapa[fornecimento]
                print(f"Nova fatura -> Fornecimento: {fornecimento} -> Codigo: {codigo_atual}")

        codigo = codigo_atual

        if codigo:
            packet = BytesIO()
            largura = float(page.mediabox.width)
            altura = float(page.mediabox.height)

            can = canvas.Canvas(packet, pagesize=(largura, altura))
            can.setFont("Helvetica-Bold", 12)
            can.drawString(largura - 150, altura - 30, f"COD: {codigo}")
            can.save()

            packet.seek(0)
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])

        paginas.append((codigo or "99999", page))

    paginas.sort(
        key=lambda x: [
            int(s) if s.isdigit() else s
            for s in re.findall(r"\d+|[A-Z]+", x[0])
        ]
    )

    writer = PdfWriter()
    for _, page_ordered in paginas:
        writer.add_page(page_ordered)

    with open(pdf_saida, "wb") as f:
        writer.write(f)

def extrair_dados_enel(pdf):
    with open(CSV_ENEL, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina", "Instalacao", "Referencia", "Consumo_kWh", "Total_Pagar", "IR_1_20"])

    reader = PdfReader(pdf)

    with open(CSV_ENEL, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")

        for i, page in enumerate(reader.pages):
            texto = page.extract_text() or ""
            if not pagina_eh_fatura(texto):
                continue

            texto_norm = normalizar(texto)
            instalacao = extrair_instalacao(texto_norm)

            writer.writerow([
                i+1,
                instalacao,
                extrair_referencia(texto_norm, instalacao),
                extrair_consumo_enel(texto),
                extrair_total(texto_norm),
                extrair_ir(texto_norm)
            ])


@app.post("/baixar-enel")
def baixar_enel():

    def gerador():
        erros = []
        processados = 0

        yield "STATUS|Conectando ao servidor de e-mail da Enel...\n"
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL_ENEL, SENHA_ENEL)
        mail.select("INBOX")

        status, mensagens = mail.search(None, "ALL")

        if not mensagens[0]:
            yield "CSV|Nenhum e-mail encontrado."
            return

        lista_emails = mensagens[0].split()
        yield f"STATUS|Encontrados {len(lista_emails)} e-mails. Baixando anexos...\n"

        sem_senha = []

        for num in lista_emails:
            status, dados = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(dados[0][1])

            remetente = msg.get("From", "").lower()
            assunto   = msg.get("Subject", "(sem assunto)")

            print(f"\n{'='*60}")
            print(f"[EMAIL #{num.decode()}]")
            print(f"  De:      {remetente}")
            print(f"  Assunto: {assunto}")

            partes_encontradas = 0
            for parte in msg.walk():
                content_type     = parte.get_content_type()
                filename_raw     = parte.get_filename()          # pode estar encodado
                filename_decoded = decodificar(filename_raw) if filename_raw else None
                disposition      = str(parte.get("Content-Disposition", ""))
                partes_encontradas += 1

                print(f"  PARTE {partes_encontradas}: tipo={content_type} | filename_raw={filename_raw} | filename_decoded={filename_decoded}")

                if not filename_decoded:
                    print(f"    -> Ignorado: sem filename")
                    continue

                if not filename_decoded.lower().endswith(".pdf"):
                    print(f"    -> Ignorado: nao e PDF ({filename_decoded})")
                    continue

                nome = f"enel_{num.decode()}_{filename_decoded}"
                caminho = os.path.join(PASTA_ENEL, nome)
                caminho_sem_senha = os.path.join(PASTA_ENEL_SEM_SENHA, nome)

                print(f"    -> PDF detectado: {nome}")
                print(f"    -> Ja existe em enel_pdf?   {os.path.exists(caminho)}")
                print(f"    -> Ja existe sem senha?      {os.path.exists(caminho_sem_senha)}")

                if not os.path.exists(caminho):
                    with open(caminho, "wb") as f:
                        f.write(parte.get_payload(decode=True))
                    print(f"    -> Baixado para enel_pdf")

                if not os.path.exists(caminho_sem_senha):
                    print(f"    -> Tentando remover senha: '{SENHA_ENEL_PDF}'")
                    ok = tentar_remover_senha(caminho, caminho_sem_senha, [SENHA_ENEL_PDF])
                    if not ok:
                        erros.append(f"Falha ao remover senha do arquivo: {nome}")
                        print(f"    -> FALHA ao remover senha")
                        continue
                    print(f"    -> Senha removida com sucesso")

                if os.path.exists(caminho_sem_senha):
                    sem_senha.append(caminho_sem_senha)
                    processados += 1
                    print(f"    -> Adicionado a lista sem_senha (total acumulado: {len(sem_senha)})")
                else:
                    print(f"    -> ERRO: arquivo sem senha nao existe apos processamento")

            print(f"  [FIM EMAIL #{num.decode()}] Partes analisadas: {partes_encontradas}")

        mail.logout()

        print(f"\n{'='*60}")
        print(f"[RESUMO] PDFs na lista sem_senha: {len(sem_senha)}")
        print(f"[RESUMO] Arquivos em PASTA_ENEL_SEM_SENHA:")
        for arq in os.listdir(PASTA_ENEL_SEM_SENHA):
            print(f"  - {arq}")
        print(f"{'='*60}\n")

        yield "STATUS|Juntando os PDFs em um unico arquivo...\n"

        for arquivo in os.listdir(PASTA_ENEL_SEM_SENHA):
            caminho = os.path.join(PASTA_ENEL_SEM_SENHA, arquivo)
            if caminho.lower().endswith(".pdf") and caminho not in sem_senha:
                sem_senha.append(caminho)
                print(f"[JUNTAR] Adicionado da pasta (execucao anterior): {arquivo}")

        print(f"[JUNTAR] Total de PDFs para juntar: {len(sem_senha)}")

        pdf_unico = os.path.join(BASE_DIR, "enel_completo.pdf")
        sucesso = juntar_pdfs(sem_senha, pdf_unico)

        if not sucesso:
            yield "CSV|Nenhum PDF ENEL encontrado"
            return

        yield "STATUS|Filtrando paginas que nao sao faturas (cartas, avisos)...\n"

        try:
            filtrar_pdf_enel(pdf_unico, PDF_ENEL_FILTRADO)
        except ValueError as e:
            yield f"CSV|{e}"
            return

        yield "STATUS|Analisando layout e escrevendo os codigos de instalacao...\n"
        mapa = carregar_mapa_instalacao_codigo()
        escrever_codigo_e_ordenar(PDF_ENEL_FILTRADO, PDF_ENEL_COM_CODIGO, mapa)

        yield "STATUS|Extraindo valores, consumos e gerando a planilha...\n"
        extrair_dados_enel(PDF_ENEL_COM_CODIGO)

        with open(CSV_ENEL, "r", encoding="utf-8-sig") as f:
            conteudo = f.read()

        resumo = f"{processados} faturas processadas com sucesso"
        if erros:
            resumo += f"\n{len(erros)} faturas com erro"

        conteudo_final = resumo + "\n\n"
        if erros:
            conteudo_final += "\n".join(erros) + "\n\n"
        conteudo_final += conteudo

        yield "STATUS|Finalizando e enviando dados para a tela...\n"
        yield f"CSV|{conteudo_final}"

    return StreamingResponse(gerador(), media_type="text/plain; charset=utf-8")


# ================= ENDPOINT SABESP =================

@app.post("/baixar-sabesp")
def baixar_sabesp():

    def gerador():
        erros = []
        processados = 0

        yield "STATUS|Conectando ao servidor de e-mail da Sabesp...\n"
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
        mail.login(EMAIL_SABESP, SENHA_SABESP)
        mail.select("INBOX")

        status, mensagens = mail.search(None, "ALL")
        lista_emails = mensagens[0].split()

        yield f"STATUS|Encontrados {len(lista_emails)} e-mails. Baixando anexos...\n"
        sem_senha = []

        for num in lista_emails:
            status, dados = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(dados[0][1])
            remetente_email = parseaddr(msg.get("From"))[1].lower()

            if remetente_email in REMETENTES_SABESP:
                for parte in msg.walk():
                    filename_raw     = parte.get_filename()
                    filename_decoded = decodificar(filename_raw) if filename_raw else None

                    if not filename_decoded or not filename_decoded.lower().endswith(".pdf"):
                        continue

                    nome = f"sabesp_{num.decode()}_{filename_decoded}"
                    caminho = os.path.join(PASTA_SABESP, nome)
                    caminho_sem_senha = os.path.join(PASTA_SABESP_SEM_SENHA, nome)

                    if not os.path.exists(caminho):
                        with open(caminho, "wb") as f:
                            f.write(parte.get_payload(decode=True))

                    if not os.path.exists(caminho_sem_senha):
                        ok = tentar_remover_senha(caminho, caminho_sem_senha, SENHAS_SABESP)
                        if not ok:
                            erros.append(f"Falha ao remover senha do arquivo: {nome}")
                            continue

                    if os.path.exists(caminho_sem_senha):
                        sem_senha.append(caminho_sem_senha)
                        processados += 1

        mail.logout()

        yield "STATUS|Juntando os PDFs em um unico arquivo...\n"
        for arquivo in os.listdir(PASTA_SABESP_SEM_SENHA):
            caminho = os.path.join(PASTA_SABESP_SEM_SENHA, arquivo)
            if caminho.lower().endswith(".pdf") and caminho not in sem_senha:
                sem_senha.append(caminho)

        sucesso = juntar_pdfs(sem_senha, PDF_SABESP_COMPLETO)

        if not sucesso:
            yield "CSV|Nenhum PDF valido encontrado"
            return

        PDF_SABESP_SEM_DUP = os.path.join(BASE_DIR, "sabesp_sem_duplicatas.pdf")
        yield "STATUS|Verificando e removendo faturas duplicadas...\n"
        try:
            duplicatas = deduplificar_pdf_sabesp(PDF_SABESP_COMPLETO, PDF_SABESP_SEM_DUP)
            if duplicatas > 0:
                yield f"STATUS|{duplicatas} fatura(s) duplicada(s) removida(s).\n"
        except ValueError as e:
            yield f"CSV|{e}"
            return

        yield "STATUS|Analisando layout e escrevendo os codigos de instalacao...\n"
        mapa = carregar_mapa_fornecimento_codigo_sabesp()
        escrever_codigo_e_ordenar_sabesp(PDF_SABESP_SEM_DUP, PDF_SABESP_COM_CODIGO, mapa)

        yield "STATUS|Extraindo valores, consumos e gerando a planilha...\n"
        extrair_dados_sabesp(PDF_SABESP_COM_CODIGO)

        with open(CSV_SABESP, "r", encoding="utf-8-sig") as f:
            conteudo = f.read()

        resumo = f"{processados} faturas processadas com sucesso"
        if erros:
            resumo += f"\n{len(erros)} faturas com erro"

        conteudo_final = resumo + "\n\n"
        if erros:
            conteudo_final += "\n".join(erros) + "\n\n"
        conteudo_final += conteudo

        yield "STATUS|Finalizando e enviando dados para a tela...\n"
        yield f"CSV|{conteudo_final}"

    return StreamingResponse(gerador(), media_type="text/plain; charset=utf-8")

# ================= FRONT =================
app.mount("/", StaticFiles(directory=FRONT_DIR, html=True), name="frontend")
