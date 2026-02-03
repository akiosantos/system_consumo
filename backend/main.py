from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import imaplib
import email
import os
from email.header import decode_header
from pypdf import PdfReader, PdfWriter
import pdfplumber
import re
import csv

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

IMAP_SERVER = "mail.barueri.sp.gov.br"

BASE_DIR = r"U:\BackupContabilidade\Custos\0 - Enel, Sabesp e Telefônica - Lucas\system\backend"

# ===== SABESP =====
EMAIL_SABESP = "seu_email"
SENHA_SABESP = "sua senha"
REMETENTE_SABESP = "fatura_sabesp@sabesp.com.br"

PASTA_SABESP = os.path.join(BASE_DIR, "sabesp_pdf")
PASTA_SABESP_SEM_SENHA = os.path.join(BASE_DIR, "sabesp_pdf_sem_senha")
CSV_SABESP = os.path.join(BASE_DIR, "sabesp_consolidado.csv")

SENHAS_SABESP = ["465", "MIG"]

# ===== ENEL =====
EMAIL_ENEL = "seu_email"
SENHA_ENEL = "sua senha"
REMETENTE_ENEL = "brasil.enel.com"

PASTA_ENEL = os.path.join(BASE_DIR, "enel_pdf")
PASTA_ENEL_SEM_SENHA = os.path.join(BASE_DIR, "enel_pdf_sem_senha")
PDF_ENEL_FILTRADO = os.path.join(BASE_DIR, "enel_filtrado.pdf")
CSV_ENEL = os.path.join(BASE_DIR, "enel_consolidado.csv")

SENHA_ENEL_PDF = "46523"

os.makedirs(PASTA_SABESP, exist_ok=True)
os.makedirs(PASTA_SABESP_SEM_SENHA, exist_ok=True)
os.makedirs(PASTA_ENEL, exist_ok=True)
os.makedirs(PASTA_ENEL_SEM_SENHA, exist_ok=True)

# ===== UTIL =====
def decodificar(texto):
    partes = decode_header(texto)
    resultado = ""
    for parte, enc in partes:
        if isinstance(parte, bytes):
            resultado += parte.decode(enc or "utf-8", errors="ignore")
        else:
            resultado += parte
    return resultado

# ===== REMOVE SENHA =====
def tentar_remover_senha(caminho_entrada, caminho_saida, senhas):
    reader = PdfReader(caminho_entrada)
    senha_ok = None

    if reader.is_encrypted:
        for senha in senhas:
            try:
                if reader.decrypt(senha) != 0:
                    senha_ok = senha
                    break
            except:
                continue

        if not senha_ok:
            return False

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with open(caminho_saida, "wb") as f:
        writer.write(f)

    return True

# ===== SABESP EXTRAÇÃO =====
def extrair_fornecimento(texto):
    m = re.search(r'\b(\d{9,16})\b', texto)
    return m.group(1) if m else None

def extrair_vencimento(texto):
    m = re.search(r'VENCIMENTO:\s*(\d{2}/\d{2}/\d{4})', texto)
    return m.group(1) if m else None

def extrair_valor_total(texto):
    texto_limpo = texto.replace("*", "")
    m = re.search(r"R\$\s*([\d.,]+)", texto_limpo)
    return m.group(1) if m else None

def extrair_retencao(texto):
    m = re.search(r'Retenção:\s*4,8%\s*([\d.,]+)', texto)
    return m.group(1) if m else None

def extrair_consumo(texto):
    texto = re.sub(r"\s+", " ", texto)

    m = re.search(
        r"\d{2}/\d{2}/\d{2}\s+\d{2}/\d{2}/\d{2}\s+(\d{1,4})\s+\d{1,4}",
        texto
    )
    if m:
        return m.group(1)

    m = re.search(
        r"\d{2}/\d{2}/\d{2}\s+\d{1,6}\s+(\d{1,4})\s+\d{1,4}",
        texto
    )
    if m:
        return m.group(1)

    return None

def pagina_principal(texto):
    t = texto.upper()
    return "FATURAMENTO" in t and "FORNECIMENTO" in t

def extrair_dados_sabesp():
    registros = []
    fornecimentos_lidos = set()

    for arquivo in os.listdir(PASTA_SABESP_SEM_SENHA):
        if not arquivo.lower().endswith(".pdf"):
            continue

        caminho_pdf = os.path.join(PASTA_SABESP_SEM_SENHA, arquivo)

        with pdfplumber.open(caminho_pdf) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                texto = page.extract_text() or ""

                if not pagina_principal(texto):
                    continue

                fornecimento = extrair_fornecimento(texto)
                if not fornecimento:
                    continue

                if fornecimento in fornecimentos_lidos:
                    continue

                vencimento = extrair_vencimento(texto)
                valor_total = extrair_valor_total(texto)
                retencao = extrair_retencao(texto)
                consumo = extrair_consumo(texto)

                registros.append([
                    arquivo,
                    fornecimento,
                    vencimento,
                    consumo,
                    valor_total,
                    retencao
                ])

                fornecimentos_lidos.add(fornecimento)

    with open(CSV_SABESP, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "Arquivo",
            "Fornecimento",
            "Vencimento",
            "Consumo_M3",
            "Valor_Total",
            "Retencao_4_8"
        ])
        writer.writerows(registros)

    return registros

# ===== ENEL =====
def tem_valor_monetario(texto):
    if not texto:
        return False
    return bool(re.search(r"r\$\s*\d+[.,]\d{2}", texto.lower()))

def pagina_interessa(texto):
    if not texto:
        return False

    t = texto.lower()
    tem_instalacao = "instala" in t or "uc" in t
    tem_vencimento = "vencimento" in t
    tem_valor = tem_valor_monetario(t)

    return sum([tem_instalacao, tem_vencimento, tem_valor]) >= 2

def pagina_eh_fatura(texto):
    t = texto.lower()
    pontos = 0
    if "instala" in t or "uc" in t:
        pontos += 1
    if "vencimento" in t:
        pontos += 1
    if re.search(r"r\$\s*\d", t):
        pontos += 1
    return pontos >= 2


def filtrar_pdf_enel(pdf_entrada, pdf_saida):
    reader = PdfReader(pdf_entrada)
    writer = PdfWriter()

    manteve_alguma = False

    for i, page in enumerate(reader.pages):
        texto = page.extract_text()
        if texto and pagina_eh_fatura(texto):
            writer.add_page(page)
            manteve_alguma = True
            print(f"✅ Página {i+1} mantida")
        else:
            print(f"❌ Página {i+1} descartada")

    # MESMO QUE NÃO TENHA PÁGINA, cria o arquivo
    with open(pdf_saida, "wb") as f:
        if manteve_alguma:
            writer.write(f)
        else:
            # cria PDF vazio para evitar erro depois
            PdfWriter().write(f)



def normalizar(texto):
    return re.sub(r"\s+", " ", texto.lower()) if texto else ""


def extrair_instalacao(texto):
    m = re.search(r"\b(\d{8,12})\s*/\s*(\d{8,13})\b", texto)
    if m:
        return m.group(1)

    padroes = [
        r"instala[çc][aã]o[^0-9]{0,10}(\d[\d\s]{5,15})",
        r"\buc\b[^0-9]{0,10}(\d[\d\s]{5,15})",
        r"unidade\s+consumidora[^0-9]{0,10}(\d[\d\s]{5,15})",
        r"contrato[^0-9]{0,10}(\d[\d\s]{5,15})",
    ]

    for p in padroes:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            return re.sub(r"\D", "", m.group(1))

    nums = re.findall(r"\b\d{8,12}\b", texto)
    return max(nums, key=len) if nums else ""


def extrair_referencia(texto, instalacao):
    t = re.sub(r"\s+", " ", texto)
    pos = t.find(instalacao)
    area = t[pos:pos+500] if pos != -1 else t
    area = re.sub(r"\b\d{2}/\d{2}/\d{4}\b", "", area)

    m = re.search(r"\b(0[1-9]|1[0-2])/[0-9]{4}\b", area)
    return m.group(0) if m else ""


def extrair_total(texto):
    padroes = [
        r"total\s+a\s+pagar\s*r?\$?\s*([\d.,]+)",
        r"valor\s+total\s*r?\$?\s*([\d.,]+)",
        r"total\s+da\s+fatura\s*r?\$?\s*([\d.,]+)",
    ]
    for p in padroes:
        m = re.search(p, texto, re.IGNORECASE)
        if m:
            return m.group(1)

    m = re.search(r"r\$\s*([\d.,]+)", texto.lower())
    return m.group(1) if m else ""


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


def extrair_dados_enel(pdf_filtrado):
    registros = []

    # cria CSV com cabeçalho sempre
    with open(CSV_ENEL, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow([
            "Pagina",
            "Instalacao",
            "Referencia",
            "Consumo_kWh",
            "Total_Pagar",
            "IR_1_20"
        ])

    if not os.path.exists(pdf_filtrado):
        print("⚠️ PDF filtrado não existe.")
        return registros

    reader = PdfReader(pdf_filtrado)

    if len(reader.pages) == 0:
        print("⚠️ PDF filtrado sem páginas.")
        return registros

    total_faturas = 0

    with open(CSV_ENEL, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")

        for i, page in enumerate(reader.pages):
            texto_original = page.extract_text()
            if not texto_original:
                continue

            if not pagina_eh_fatura(texto_original):
                continue

            texto = normalizar(texto_original)

            instalacao = extrair_instalacao(texto)
            referencia = extrair_referencia(texto, instalacao)
            total = extrair_total(texto)
            ir = extrair_ir(texto)
            consumo = extrair_consumo_enel(texto_original)

            writer.writerow([
                i + 1,
                instalacao,
                referencia,
                consumo,
                total,
                ir
            ])

            total_faturas += 1
            print(f"Página {i+1} OK")

    return registros



# ===== ENDPOINT SABESP =====
@app.post("/baixar-sabesp")
def baixar_sabesp():
    baixados = []
    desbloqueados = []
    falha_senha = []

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    mail.login(EMAIL_SABESP, SENHA_SABESP)
    mail.select("INBOX")

    status, mensagens = mail.search(None, "ALL")

    for num in mensagens[0].split():
        status, dados = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(dados[0][1])

        if REMETENTE_SABESP in msg.get("From", "").lower():
            for parte in msg.walk():
                if parte.get_content_disposition() == "attachment":
                    nome_original = decodificar(parte.get_filename() or "fatura.pdf")
                    nome = f"sabesp_{num.decode()}_{nome_original}"

                    caminho = os.path.join(PASTA_SABESP, nome)

                    if not os.path.exists(caminho):
                        with open(caminho, "wb") as f:
                            f.write(parte.get_payload(decode=True))
                        baixados.append(nome)

                        destino = os.path.join(PASTA_SABESP_SEM_SENHA, nome)
                        ok = tentar_remover_senha(caminho, destino, SENHAS_SABESP)

                        if ok:
                            desbloqueados.append(nome)
                        else:
                            falha_senha.append(nome)

    mail.logout()
    extrair_dados_sabesp()

    return FileResponse(
        path=CSV_SABESP,
        filename="sabesp_consolidado.csv",
        media_type="text/csv"
    )

# ===== ENDPOINT ENEL =====
# ===== ENDPOINT ENEL =====
@app.post("/baixar-enel")
def baixar_enel():
    baixados = []
    sem_senha_pdfs = []

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    mail.login(EMAIL_ENEL, SENHA_ENEL)
    mail.select("INBOX")

    status, mensagens = mail.search(None, "ALL")

    for num in mensagens[0].split():
        status, dados = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(dados[0][1])

        remetente = msg.get("From", "").lower()
        print("DEBUG remetente:", remetente)

        if REMETENTE_ENEL in remetente:
            for parte in msg.walk():
                content_type = parte.get_content_type()

                if content_type == "application/pdf":
                    nome_original = decodificar(parte.get_filename() or f"fatura_enel_{num.decode()}.pdf")
                    nome = f"enel_{num.decode()}_{nome_original}"

                    caminho = os.path.join(PASTA_ENEL, nome)

                    if not os.path.exists(caminho):
                        with open(caminho, "wb") as f:
                            f.write(parte.get_payload(decode=True))
                        print("✅ Baixado:", nome)

                        sem_senha = os.path.join(PASTA_ENEL_SEM_SENHA, nome)
                        ok = tentar_remover_senha(caminho, sem_senha, [SENHA_ENEL_PDF])

                        if ok:
                            sem_senha_pdfs.append(sem_senha)
                            baixados.append(nome)

    mail.logout()

    if sem_senha_pdfs:
        pdf_unico = os.path.join(BASE_DIR, "enel_completo.pdf")
        writer_pdf = PdfWriter()

        for pdf_path in sem_senha_pdfs:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer_pdf.add_page(page)

        with open(pdf_unico, "wb") as f:
            writer_pdf.write(f)

        filtrar_pdf_enel(pdf_unico, PDF_ENEL_FILTRADO)
        extrair_dados_enel(PDF_ENEL_FILTRADO)
    else:
        extrair_dados_enel(PDF_ENEL_FILTRADO)

    return FileResponse(
        path=CSV_ENEL,
        filename="enel_consolidado.csv",
        media_type="text/csv"
    )


