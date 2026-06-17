import base64
import logging
import os
import re
import xmlrpc.client
import requests
from fastapi import FastAPI, Request, BackgroundTasks

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

app = FastAPI()

# =========================
# CONFIG
# =========================

ODOO_URL = "https://yep.yafrel.com"
ODOO_DB = "yafrel-education-platform"

ODOO_USER_ID = int(os.getenv("ODOO_USER_ID", "2"))
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_USER_EMAIL = "yafrelservices@yafrel.com"

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

DELETE_ATTACHMENTS = True

# carpeta FIJA (evita caos)
BASE_FOLDER_NAME = "Yafrel Medical Care"

# =========================
# VALIDACION
# =========================

if not all([ODOO_PASSWORD, AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
    raise RuntimeError("Faltan variables de entorno")

# =========================
# HELPERS
# =========================

def limpiar(nombre):
    return re.sub(r'[<>:"/\\|?*]', "_", nombre).strip()

def token():
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    r = requests.post(url, data={
        "grant_type": "client_credentials",
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    })
    r.raise_for_status()
    return r.json()["access_token"]

def headers():
    return {"Authorization": f"Bearer {token()}"}

# =========================
# ROOT FIX (NO DUPLICAR)
# =========================

def get_base_folder():
    url = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"
    r = requests.get(url, headers=headers())
    r.raise_for_status()

    for f in r.json().get("value", []):
        if f["name"] == BASE_FOLDER_NAME:
            return f["id"]

    r = requests.post(url, json={
        "name": BASE_FOLDER_NAME,
        "folder": {}
    }, headers=headers)

    r.raise_for_status()
    return r.json()["id"]

# =========================
# PROCESS
# =========================

def process(payload):

    applicant_id = payload.get("id")
    name = limpiar(payload.get("display_name") or f"Candidato_{applicant_id}")

    logging.info(f"Procesando {name} ({applicant_id})")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    base_id = get_base_folder()

    # ODOO check
    exists = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "search_count",
        [[["name", "=", f"Expediente - {name}"]]]
    )

    if exists:
        logging.info("Ya existe en ODOO")
        return

    # CREATE folder (SIN rename)
    url = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{base_id}/children"

    r = requests.post(url, json={
        "name": f"{applicant_id}_{name}",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail"
    }, headers=headers())

    # si falla por duplicado → NO crear otra
    if r.status_code not in (200, 201):
        logging.error(f"Folder error: {r.text}")
        return

    folder_id = r.json()["id"]

    # CREATE SHARE LINK
    link = requests.post(
        f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink",
        json={"type": "view", "scope": "organization"},
        headers=headers()
    )

    link.raise_for_status()

    url_link = link.json()["link"]["webUrl"]

    # ODOO CREATE (VISIBLE GLOBAL)
    doc_id = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "create",
        [{
            "name": f"Expediente - {name}",
            "type": "url",
            "url": url_link
        }]
    )

    logging.info(f"ODOO OK {doc_id}")

    # DELETE attachments (solo si existe flujo real)
    if DELETE_ATTACHMENTS:
        att = payload.get("attachment_ids") or []
        if att:
            models.execute_kw(
                ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
                "ir.attachment", "unlink",
                [att]
            )

# =========================
# WEBHOOK
# =========================

@app.post("/webhook")
async def webhook(req: Request, bg: BackgroundTasks):
    data = await req.json()
    if data.get("id"):
        bg.add_task(process, data)
    return {"ok": True}

@app.get("/")
def health():
    return {"status": "ok"}
