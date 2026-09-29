"""
Disparador de Emails - Interface Web
Flask server que expoe APIs para o frontend HTML
"""

import os, re, base64, logging, threading, io
from pathlib import Path
from collections import defaultdict
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import openpyxl
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Caminhos: sempre relativos ao app.py, independente de onde for executado
PROJETO_DIR = Path(__file__).resolve().parent
BASE_DIR    = PROJETO_DIR.parent

CREDENTIALS = str(PROJETO_DIR / "credentials.json")
TOKEN_FILE  = str(PROJETO_DIR / "token.json")
SCOPES      = ["https://www.googleapis.com/auth/gmail.send"]
PADRAO_PDF  = re.compile(r"^(.+)\s+USUARIO\s+(\d+)\.pdf$", re.IGNORECASE)

PASTA_PDFS = BASE_DIR / "PDFs"
PASTA_XLSX = BASE_DIR / "Clientes" / "clientes.xlsx"

print(f"[INFO] Raiz : {BASE_DIR}")
print(f"[INFO] PDFs : {PASTA_PDFS}")
print(f"[INFO] XLSX : {PASTA_XLSX}")

app = Flask(__name__, static_folder="static")
CORS(app)

# Estado em memoria — planilha lida como bytes, sem depender de caminho fixo
memoria = {"xlsx_bytes": None, "pdfs_dir": None}

estado_disparo = {
    "rodando": False, "total": 0, "enviados": 0,
    "erros": 0, "sem_pdf": 0, "log": [], "concluido": False
}

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

def autenticar_gmail():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
    return build("gmail", "v1", credentials=creds)

def mapear_pdfs(pasta):
    mapa = defaultdict(list)
    for arq in sorted(Path(pasta).glob("*.pdf")):
        m = PADRAO_PDF.match(arq.name)
        if m:
            mapa[m.group(2).strip()].append((m.group(1).strip(), arq))
    return mapa

def ler_planilha_bytes(xlsx_bytes, data_override=None, cc_override=None):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

    # Tenta ler da aba Configuracoes; se não existir, usa os valores da Etapa 1
    data_leilao = data_override or ""
    email_cc    = cc_override or ""
    if "Configuracoes" in wb.sheetnames:
        ws_cfg   = wb["Configuracoes"]
        val_data = ws_cfg["B4"].value
        data_leilao = val_data.strftime("%d/%m/%Y") if hasattr(val_data, "strftime") else str(val_data or "").strip()
        email_cc    = str(ws_cfg["B5"].value or "").strip()

    # Aceita qualquer aba: Clientes, clientes, Sheet1, a primeira disponível
    aba_clientes = None
    for nome in ["Clientes", "clientes", "CLIENTES"]:
        if nome in wb.sheetnames:
            aba_clientes = wb[nome]
            break
    if aba_clientes is None:
        aba_clientes = wb.active  # usa a primeira aba disponível

    clientes = {}
    for row in aba_clientes.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        id_c = str(row[0]).strip().zfill(6)
        clientes[id_c] = {
            "nome":  str(row[1] or "").strip(),
            "email": str(row[2] or "").strip(),
            "lotes": str(row[3] or "").strip(),
        }
    return clientes, data_leilao, email_cc

def montar_assunto(nome, id_c, eventos, data, lotes):
    evs = sorted(set(eventos))
    parte_lotes = f"Lotes {lotes} - " if lotes else ""
    return f"Documentos Leilao {data} - {parte_lotes}{', '.join(evs)} - {nome} - {id_c}"

def montar_email(nome, email, assunto, corpo_html, pdfs, email_cc):
    msg = MIMEMultipart("mixed")
    msg["To"]      = f"{nome} <{email}>"
    if email_cc:
        msg["Cc"] = email_cc
    msg["Subject"] = assunto
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(corpo_html, "html", "utf-8"))
    msg.attach(alt)
    for _, path in pdfs:
        with open(path, "rb") as f:
            parte = MIMEApplication(f.read(), _subtype="pdf")
            parte.add_header("Content-Disposition", "attachment", filename=path.name)
            msg.attach(parte)
    return msg

def enviar(servico, msg, email_cc=None):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    body = {"raw": raw}
    servico.users().messages().send(userId="me", body=body).execute()

@app.route("/api/salvar-config", methods=["POST"])
def salvar_config():
    data = request.json
    memoria["config_data"] = data.get("data", "")
    memoria["config_cc"]   = data.get("cc", "")
    return jsonify({"ok": True})

@app.route("/api/auth-status")
def auth_status():
    ok = os.path.exists(TOKEN_FILE) and os.path.exists(CREDENTIALS)
    return jsonify({"autenticado": ok})

@app.route("/api/limpar-planilha", methods=["POST"])
def limpar_planilha():
    memoria["xlsx_bytes"] = None
    return jsonify({"ok": True})

@app.route("/api/debug")
def debug():
    return jsonify({
        "PROJETO_DIR": str(PROJETO_DIR),
        "BASE_DIR":    str(BASE_DIR),
        "PASTA_PDFS":  str(PASTA_PDFS),
        "PASTA_XLSX":  str(PASTA_XLSX),
    })

@app.route("/api/upload-planilha", methods=["POST"])
def upload_planilha():
    try:
        arquivo = request.files["planilha"]
        PASTA_XLSX.parent.mkdir(parents=True, exist_ok=True)
        arquivo.save(str(PASTA_XLSX))                  # salva direto na pasta Clientes/
        xlsx_bytes = PASTA_XLSX.read_bytes()
        memoria["xlsx_bytes"] = xlsx_bytes
        data_cfg = memoria.get("config_data", "")
        cc_cfg   = memoria.get("config_cc", "")
        clientes, data, cc = ler_planilha_bytes(xlsx_bytes, data_cfg, cc_cfg)
        preview = [{"id": k, **v} for k, v in list(clientes.items())[:5]]
        return jsonify({"ok": True, "total": len(clientes), "data": data, "cc": cc, "preview": preview})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)}), 400

@app.route("/api/upload-pdfs", methods=["POST"])
def upload_pdfs():
    try:
        PASTA_PDFS.mkdir(parents=True, exist_ok=True)
        memoria["pdfs_dir"] = str(PASTA_PDFS)
        salvos, invalidos = [], []
        for arq in request.files.getlist("pdfs"):
            if PADRAO_PDF.match(arq.filename):
                arq.save(str(PASTA_PDFS / arq.filename))
                salvos.append(arq.filename)
            else:
                invalidos.append(arq.filename)
        return jsonify({"ok": True, "salvos": len(salvos), "invalidos": invalidos, "nomes": salvos})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)}), 400

@app.route("/api/preview", methods=["POST"])
def preview():
    try:
        if not memoria["xlsx_bytes"]:
            return jsonify({"ok": False, "erro": "Planilha nao carregada. Volte para a Etapa 2."}), 400
        pdfs_dir = memoria["pdfs_dir"] or str(PASTA_PDFS)
        clientes, data, cc = ler_planilha_bytes(memoria["xlsx_bytes"])
        mapa = mapear_pdfs(pdfs_dir)
        previews = []
        for id_c, dados in clientes.items():
            pdfs = mapa.get(id_c, [])
            if not pdfs:
                continue
            eventos = [e for e, _ in pdfs]
            assunto = montar_assunto(dados["nome"], id_c, eventos, data, dados["lotes"])
            previews.append({
                "nome":    dados["nome"],
                "email":   dados["email"],
                "assunto": assunto,
                "anexos":  [p.name for _, p in pdfs],
                "lotes":   dados["lotes"],
            })
        return jsonify({"ok": True, "previews": previews, "total": len(previews), "cc": cc})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)}), 400

@app.route("/api/disparar", methods=["POST"])
def disparar():
    global estado_disparo
    if estado_disparo["rodando"]:
        return jsonify({"ok": False, "erro": "Disparo ja em andamento"}), 400
    if not memoria["xlsx_bytes"]:
        return jsonify({"ok": False, "erro": "Planilha nao carregada. Volte para a Etapa 2."}), 400
    corpo_html = request.json.get("corpo", "")

    def run():
        global estado_disparo
        estado_disparo = {"rodando": True, "total": 0, "enviados": 0,
                          "erros": 0, "sem_pdf": 0, "log": [], "concluido": False}
        try:
            servico  = autenticar_gmail()
            pdfs_dir = memoria["pdfs_dir"] or str(PASTA_PDFS)
            data_cfg = memoria.get("config_data", "")
            cc_cfg   = memoria.get("config_cc", "")
            clientes, data, cc = ler_planilha_bytes(memoria["xlsx_bytes"], data_cfg, cc_cfg)
            # Se o CC ainda estiver vazio, usa o das configurações da etapa 1
            if not cc:
                cc = cc_cfg
            mapa = mapear_pdfs(pdfs_dir)
            estado_disparo["total"] = len(clientes)
            estado_disparo["log"].append({"tipo": "info", "msg": f"Iniciando disparo para {len(clientes)} cliente(s)..."})
            for id_c, dados in clientes.items():
                pdfs = mapa.get(id_c)
                if not pdfs:
                    estado_disparo["sem_pdf"] += 1
                    estado_disparo["log"].append({"tipo": "warn", "msg": f"Sem PDF: {dados['nome']} ({id_c})"})
                    continue
                eventos = [e for e, _ in pdfs]
                assunto = montar_assunto(dados["nome"], id_c, eventos, data, dados["lotes"])
                try:
                    msg = montar_email(dados["nome"], dados["email"], assunto, corpo_html, pdfs, cc)
                    enviar(servico, msg)
                    estado_disparo["enviados"] += 1
                    estado_disparo["log"].append({"tipo": "ok", "msg": f"OK {dados['nome']} - {dados['email']} - {len(pdfs)} anexo(s)"})
                except Exception as e:
                    estado_disparo["erros"] += 1
                    estado_disparo["log"].append({"tipo": "erro", "msg": f"ERRO {dados['nome']}: {str(e)}"})
            estado_disparo["log"].append({"tipo": "info", "msg": f"Concluido: {estado_disparo['enviados']} enviados | {estado_disparo['erros']} erros | {estado_disparo['sem_pdf']} sem PDF"})
        except Exception as e:
            estado_disparo["log"].append({"tipo": "erro", "msg": f"Erro fatal: {str(e)}"})
        finally:
            estado_disparo["rodando"]   = False
            estado_disparo["concluido"] = True

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"ok": True})

@app.route("/api/status")
def status():
    return jsonify(estado_disparo)

@app.route("/")
def index():
    return send_from_directory("static", "index.html")

if __name__ == "__main__":
    import webbrowser
    print("\n Disparador de Emails iniciado!")
    print(f" Acesse: http://localhost:5000\n")
    threading.Timer(1.2, lambda: webbrowser.open("http://localhost:5000")).start()
    app.run(debug=False, port=5000)
