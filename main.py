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

GRAPH = "https://graph.microsoft.com/v1.0"

BASE_FOLDER = "Recruitment"

DELETE_ATTACHMENTS = True

# =========================
# VALIDACION
# =========================

if not all([ODOO_PASSWORD, AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
    raise RuntimeError("Missing env vars")

# =========================
# HELPERS
# =========================

def clean(name):
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()

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
# GET / CREATE ROOT (RECRUITMENT)
# =========================

def get_recruitment_folder():
    url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"
    r = requests.get(url, headers=headers())
    r.raise_for_status()

    for f in r.json().get("value", []):
        if f["name"].lower() == BASE_FOLDER.lower():
            return f["id"]

    r = requests.post(url, json={
        "name": BASE_FOLDER,
        "folder": {}
    }, headers=headers())

    r.raise_for_status()
    return r.json()["id"]

# =========================
# PROCESS
# =========================

def process(payload):

    applicant_id = payload.get("id")
    name = clean(payload.get("display_name") or f"Candidate_{applicant_id}")

    logging.info(f"Processing {name} ({applicant_id})")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    root_id = get_recruitment_folder()

    # =========================
    # ODOO CHECK
    # =========================

    exists = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "search_count",
        [[["name", "=", f"Candidate Profile - {name}"]]]
    )

    if exists:
        logging.info("Already exists in Odoo")
        return

    # =========================
    # CREATE ONE DRIVE FOLDER (NO PREFIX, NO RENAMES)
    # =========================

    url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{root_id}/children"

    r = requests.post(url, json={
        "name": name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail"
    }, headers=headers())

    if r.status_code not in (200, 201):
        logging.error(f"Folder error: {r.text}")
        return

    folder_id = r.json()["id"]

    # =========================
    # CREATE SHARE LINK
    # =========================

    link = requests.post(
        f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink",
        json={"type": "view", "scope": "organization"},
        headers=headers()
    )

    link.raise_for_status()

    weburl = link.json()["link"]["webUrl"]

    # =========================
    # CREATE ODOO DOCUMENT (RECRUITMENT)
    # =========================

    doc_id = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "create",
        [{
            "name": f"Candidate Profile - {name}",
            "type": "url",
            "url": weburl
        }]
    )

    logging.info(f"Odoo document created {doc_id}")

    # =========================
    # DELETE ATTACHMENTS (OPTIONAL)
    # =========================

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
    return {"status": "running"}
