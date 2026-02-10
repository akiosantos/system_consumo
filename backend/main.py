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
import pdfplumber
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
EMAIL_SABESP = "seu_email"
SENHA_SABESP = "sua_senha"
REMETENTES_SABESP = [
    "email@remetente",
    "email@remetente1"
]

PASTA_SABESP = os.path.join(BASE_DIR, "sabesp_pdf")
PASTA_SABESP_SEM_SENHA = os.path.join(BASE_DIR, "sabesp_pdf_sem_senha")
CSV_SABESP = os.path.join(BASE_DIR, "sabesp_consolidado.csv")

PDF_SABESP_COMPLETO = os.path.join(BASE_DIR, "sabesp_completo.pdf")

SENHAS_SABESP = ["465", "MIG"]

# ===== ENEL =====
EMAIL_ENEL = "seu_email"
SENHA_ENEL = "sua_senha"
REMETENTE_ENEL = "email@remetente"

PASTA_ENEL = os.path.join(BASE_DIR, "enel_pdf")
PASTA_ENEL_SEM_SENHA = os.path.join(BASE_DIR, "enel_pdf_sem_senha")
PDF_ENEL_FILTRADO = os.path.join(BASE_DIR, "enel_filtrado.pdf")
PDF_ENEL_COM_CODIGO = os.path.join(BASE_DIR, "enel_com_codigo.pdf")
CSV_ENEL = os.path.join(BASE_DIR, "enel_consolidado.csv")

SENHA_ENEL_PDF = "46523"

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
   
    if reader.is_encrypted:
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
    for page in reader.pages:
        writer.add_page(page)

    with open(caminho_saida, "wb") as f:
        writer.write(f)

    return True

def juntar_pdfs(lista_pdfs, pdf_saida):

    writer = PdfWriter()

    for caminho in lista_pdfs:

        if not os.path.exists(caminho):
            print(f"Arquivo não encontrado: {caminho}")
            continue

        try:
            reader = PdfReader(caminho)

            for page in reader.pages:
                writer.add_page(page)

        except Exception as e:
            print(f"Erro ao abrir {caminho}: {e}")

    if len(writer.pages) == 0:
        print("Nenhum PDF válido encontrado.")
        return False

    with open(pdf_saida, "wb") as f:
        writer.write(f)

    return True


# ================= SABESP EXTRAÇÃO =================
def extrair_consumo_sabesp(texto):

    texto = re.sub(r"\s+", " ", texto)

    m = re.search(
        r"\d{2}/\d{2}/\d{2}\s+\d{2}/\d{2}/\d{2}\s+(\d{1,6})\s+\d{1,6}",
        texto
    )
    if m:
        return m.group(1)

    m = re.search(
        r"\d{2}/\d{2}/\d{2}\s+\d{1,6}\s+(\d{1,6})\s+\d{1,6}",
        texto
    )
    if m:
        return m.group(1)

    m = re.search(
        r"\d{2}/\d{2}/\d{2}.*?\d{1,6}\s+(\d{1,6})\s+\d{1,6}",
        texto
    )
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
    fornecimentos_lidos = set()

    reader = PdfReader(pdf)

    for i, page in enumerate(reader.pages):

        texto = page.extract_text() or ""

        if "FATURAMENTO" not in texto.upper():
            continue

        fornecimento = extrair_fornecimento_sabesp(texto)

        if not fornecimento or fornecimento in fornecimentos_lidos:
            continue

        vencimento = re.search(r'VENCIMENTO:\s*(\d{2}/\d{2}/\d{4})', texto)
        vencimento = vencimento.group(1) if vencimento else ""

        texto_limpo = texto.replace("*", "")

        valores = re.findall(r"R\$\s*([\d.,]+)", texto_limpo)

        if valores:

            valores_float = [
                float(v.replace(".", "").replace(",", "."))
                for v in valores
            ]

            maior = max(valores_float)

            valor = f"{maior:.2f}".replace(".", ",")

        else:
            valor = ""

        retencao = re.search(r'Retenção:\s*4,8%\s*([\d.,]+)', texto)
        retencao = retencao.group(1) if retencao else ""

        consumo = extrair_consumo_sabesp(texto)

        registros.append([
            f"Pagina {i+1}",
            fornecimento,
            vencimento,
            consumo,
            valor,
            retencao
        ])

        fornecimentos_lidos.add(fornecimento)

    with open(CSV_SABESP, "w", newline="", encoding="utf-8-sig") as f:

        writer = csv.writer(f, delimiter=";")

        writer.writerow([
            "Pagina",
            "Fornecimento",
            "Vencimento",
            "Consumo_M3",
            "Valor_Total",
            "Retencao_4_8"
        ])

        writer.writerows(registros)




# ================= PLANILHA =================
def carregar_mapa_instalacao_codigo():
    wb = load_workbook(PLANILHA_ENEL, data_only=True)
    ws = wb[ABA_ENEL]

    mapa = {}
    for linha in ws.iter_rows(min_row=2):
        codigo = linha[1].value   # B
        instalacao = linha[3].value  # D

        if codigo and instalacao:
            mapa[normalizar_instalacao(str(instalacao))] = str(codigo)

    print("Mapa instalação → código:", mapa)
    return mapa

def carregar_mapa_fornecimento_codigo_sabesp():

    wb = load_workbook(PLANILHA_SABESP, data_only=True)
    ws = wb[ABA_SABESP]

    mapa = {}

    for linha in ws.iter_rows(min_row=2):

        codigo = linha[1].value        # coluna B
        fornecimento = linha[4].value  # coluna E

        if codigo and fornecimento:

            fornecimento = str(fornecimento).strip()
            mapa[fornecimento] = str(codigo)

    print("Mapa SABESP fornecimento → código:", mapa)

    return mapa

# ================= ENEL =================
def pagina_eh_fatura(texto):
    t = texto.lower()

    # ❌ páginas que NÃO são fatura (cartas da Enel)
    palavras_excluir = [
        "faturamento a menor",
        "ausência de faturamento",
        "assunto:",
        "olá",
        "parcelado em",
        "diferença em relação",
        "consumo acumulado",
        "aqui, você pode acompanhar",
    ]

    for p in palavras_excluir:
        if p in t:
            return False

    # ✅ critérios de fatura
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

    with open(pdf_saida, "wb") as f:
        writer.write(f)

def extrair_instalacao(texto):
    m = re.search(r"\b\d{8,12}\b", texto)
    return normalizar_instalacao(m.group(0)) if m else ""

def extrair_referencia(texto, instalacao):
    if not instalacao:
        return ""

    t = re.sub(r"\s+", " ", texto)

    # garante que instalação sem zero e com zero funcionem
    instalacao = instalacao.lstrip("0")

    # procura instalação no texto
    pos = -1
    for padrao in [instalacao, instalacao.zfill(len(instalacao)+1)]:
        p = t.find(padrao)
        if p != -1:
            pos = p
            break

    area = t[pos:pos+500] if pos != -1 else t

    # remove datas completas DD/MM/AAAA
    area = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", area)

    # procura MM/AAAA
    m = re.search(r"\b(0[1-9]|1[0-2])/[0-9]{4}\b", area)

    return m.group(0) if m else ""



def extrair_total(texto):

    texto = texto.lower()

    # REGRA 1 — se existir R$***** então valor é zero
    if re.search(r"r\$\s*\*+", texto):
        return "0,00"

    # REGRA 2 — pegar valor numérico normalmente
    m = re.search(r"r\$\s*([\d.,]+)", texto)

    if m:
        return m.group(1)

    return ""


def extrair_ir(texto):
    texto = texto.lower()

    # padrão principal usado pela Enel
    m = re.search(
        r"ret\.\s*art\.\s*64\s*lei\s*9430\s*-\s*1[,\.]20%\s*(?:[\d.,]+\s*){0,3}(-?\d[\d.,]*)",
        texto
    )
    if m:
        return m.group(1).replace("-", "")

    # padrão alternativo (IRRF)
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
        m = re.search(
            r"(?:CONSUMO|USO SIST\. DISTR\.) .*?KWH\s+([\d.,]+)",
            texto
        )
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
            can = canvas.Canvas(packet)
            can.setFont("Helvetica-Bold", 12)
            can.drawString(430, 800, f"CÓD: {codigo}")
            can.save()
            packet.seek(0)
            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])

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

        # se não conseguiu extrair texto, usa vazio
        if not texto:
            texto = ""

        fornecimento = extrair_fornecimento_sabesp(texto)

        # se encontrou fornecimento novo → atualiza código atual
        if fornecimento and fornecimento in mapa:

            codigo_atual = mapa[fornecimento]

            print(f"Fornecimento encontrado: {fornecimento} → Código: {codigo_atual}")

        # usa código atual mesmo se página não tiver texto
        codigo = codigo_atual

        # escreve código na página
        if codigo:

            packet = BytesIO()

            can = canvas.Canvas(packet)

            can.setFont("Helvetica-Bold", 12)

            # posição mais alta (ajuste fino se quiser)
            can.drawString(430, 820, f"CÓD: {codigo}")

            can.save()

            packet.seek(0)

            overlay = PdfReader(packet)

            page.merge_page(overlay.pages[0])

        paginas.append((codigo or "99999", page))

    # ordenar páginas pelo código
    paginas.sort(
        key=lambda x: [
            int(s) if s.isdigit() else s
            for s in re.findall(r"\d+|[A-Z]+", x[0])
        ]
    )

    writer = PdfWriter()

    for _, page in paginas:
        writer.add_page(page)

    with open(pdf_saida, "wb") as f:
        writer.write(f)



def extrair_dados_enel(pdf):
    with open(CSV_ENEL, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["Pagina","Instalacao","Referencia","Consumo_kWh","Total_Pagar","IR_1_20"])

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

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    mail.login(EMAIL_ENEL, SENHA_ENEL)
    mail.select("INBOX")

    status, mensagens = mail.search(None, "ALL")

    sem_senha = []

    for num in mensagens[0].split():

        status, dados = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(dados[0][1])

        remetente = msg.get("From", "").lower()

        print("\n==========================")
        print("Remetente encontrado:", remetente)

        if REMETENTE_ENEL in remetente:

            print("EMAIL ENEL IDENTIFICADO")

            for parte in msg.walk():

                content_type = parte.get_content_type()
                disposition = parte.get_content_disposition()
                filename = parte.get_filename()

                print("Tipo:", content_type)
                print("Disposition:", disposition)
                print("Arquivo:", filename)

                if filename and filename.lower().endswith(".pdf"):

                    nome = f"enel_{num.decode()}_{decodificar(filename)}"

                    caminho = os.path.join(PASTA_ENEL, nome)
                    caminho_sem_senha = os.path.join(PASTA_ENEL_SEM_SENHA, nome)

                    print("Salvando:", caminho)

                    if not os.path.exists(caminho):

                        with open(caminho, "wb") as f:
                            f.write(parte.get_payload(decode=True))

                    if not os.path.exists(caminho_sem_senha):

                        ok = tentar_remover_senha(
                            caminho,
                            caminho_sem_senha,
                            [SENHA_ENEL_PDF]
                        )

                        if not ok:
                            print("Falha ao remover senha:", nome)
                            continue

                    if os.path.exists(caminho_sem_senha):

                        sem_senha.append(caminho_sem_senha)

    mail.logout()

    # incluir PDFs já existentes
    for arquivo in os.listdir(PASTA_ENEL_SEM_SENHA):

        caminho = os.path.join(PASTA_ENEL_SEM_SENHA, arquivo)

        if caminho.lower().endswith(".pdf"):

            if caminho not in sem_senha:
                sem_senha.append(caminho)

    # juntar PDFs
    pdf_unico = os.path.join(BASE_DIR, "enel_completo.pdf")

    sucesso = juntar_pdfs(sem_senha, pdf_unico)

    if not sucesso:
        return {"erro": "Nenhum PDF ENEL encontrado"}

    # filtrar
    filtrar_pdf_enel(pdf_unico, PDF_ENEL_FILTRADO)

    # inserir código
    mapa = carregar_mapa_instalacao_codigo()

    escrever_codigo_e_ordenar(
        PDF_ENEL_FILTRADO,
        PDF_ENEL_COM_CODIGO,
        mapa
    )

    # gerar CSV
    extrair_dados_enel(PDF_ENEL_COM_CODIGO)

    from fastapi.responses import Response

    with open(CSV_ENEL, "r", encoding="utf-8-sig") as f:
        conteudo = f.read()

    return Response(
        content=conteudo,
        media_type="text/plain; charset=utf-8"
    )

# ================= ENDPOINT SABESP =================
@app.post("/baixar-sabesp")
def baixar_sabesp():

    erros = []
    processados = 0

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    mail.login(EMAIL_SABESP, SENHA_SABESP)
    mail.select("INBOX")

    status, mensagens = mail.search(None, "ALL")

    sem_senha = []

    for num in mensagens[0].split():

        status, dados = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(dados[0][1])

        remetente_email = parseaddr(msg.get("From"))[1].lower()

        if remetente_email in REMETENTES_SABESP:

            print("EMAIL ACEITO PELO SISTEMA")

            for parte in msg.walk():

                print("TIPO:", parte.get_content_type())
                print("DISPOSITION:", parte.get_content_disposition())
                print("ARQUIVO:", parte.get_filename())
                print("----------------------------")

                filename = parte.get_filename()

                if filename and filename.lower().endswith(".pdf"):

                    nome = f"sabesp_{num.decode()}_{decodificar(parte.get_filename() or 'fatura.pdf')}"

                    caminho = os.path.join(PASTA_SABESP, nome)
                    caminho_sem_senha = os.path.join(PASTA_SABESP_SEM_SENHA, nome)

                    if not os.path.exists(caminho):

                        with open(caminho, "wb") as f:
                            f.write(parte.get_payload(decode=True))

                    if not os.path.exists(caminho_sem_senha):

                        ok = tentar_remover_senha(
                            caminho,
                            caminho_sem_senha,
                            SENHAS_SABESP
                        )

                        if not ok:
                            
                            erro_msg = f"❌ Falha ao remover senha do arquivo: {nome}"

                            print(erro_msg)

                            erros.append(erro_msg)

                            continue

                    if os.path.exists(caminho_sem_senha):
                        sem_senha.append(caminho_sem_senha)
                        processados += 1

    mail.logout()

    for arquivo in os.listdir(PASTA_SABESP_SEM_SENHA):

        caminho = os.path.join(PASTA_SABESP_SEM_SENHA, arquivo)

        if caminho.lower().endswith(".pdf"):

            if caminho not in sem_senha:
                sem_senha.append(caminho)

    sucesso = juntar_pdfs(sem_senha, PDF_SABESP_COMPLETO)

    if not sucesso:
        return {"erro": "Nenhum PDF válido encontrado"}

    # carregar mapa
    mapa = carregar_mapa_fornecimento_codigo_sabesp()

    # escrever código e ordenar
    escrever_codigo_e_ordenar_sabesp(
        PDF_SABESP_COMPLETO,
        PDF_SABESP_COM_CODIGO,
        mapa
    )

    # extrair dados normalmente
    extrair_dados_sabesp(PDF_SABESP_COM_CODIGO)

    from fastapi.responses import Response

    with open(CSV_SABESP, "r", encoding="utf-8-sig") as f:
        conteudo = f.read()

    resumo = f"✅ {processados} faturas processadas com sucesso"

    if erros:
        resumo += f"\n❌ {len(erros)} faturas com erro"

    conteudo_final = resumo + "\n\n"

    if erros:
        conteudo_final += "\n".join(erros) + "\n\n"

    conteudo_final += conteudo


    return Response(
        content=conteudo_final,
        media_type="text/plain; charset=utf-8"
    )


# ================= FRONT =================
app.mount("/", StaticFiles(directory=FRONT_DIR, html=True), name="frontend")
