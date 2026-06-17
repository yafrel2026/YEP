import base64
import logging
import os
import re
import xmlrpc.client
import requests
from fastapi import FastAPI, Request, BackgroundTasks

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

app = FastAPI()

# =========================================================
# CONFIG
# =========================================================

ODOO_URL = "https://yep.yafrel.com"
ODOO_DB = "yafrel-education-platform"
ODOO_USER_ID = int(os.getenv("ODOO_USER_ID", "2"))
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_USER_EMAIL = "yafrelservices@yafrel.com"

GRAPH = "https://graph.microsoft.com/v1.0"

ONEDRIVE_ROOT = "Yafrel Medical Care"
ODOO_RECRUITMENT_FOLDER = "Recruitment"

DELETE_ATTACHMENTS = True

# =========================================================
# VALIDATION
# =========================================================

if not all([ODOO_PASSWORD, AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
    raise RuntimeError("Missing env vars")

# =========================================================
# HELPERS
# =========================================================

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

# =========================================================
# ONEDRIVE ROOT (STABLE)
# =========================================================

def get_root():
    url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"
    r = requests.get(url, headers=headers())
    r.raise_for_status()

    for f in r.json().get("value", []):
        if f["name"] == ONEDRIVE_ROOT:
            return f["id"]

    r = requests.post(url, json={
        "name": ONEDRIVE_ROOT,
        "folder": {}
    }, headers=headers())

    r.raise_for_status()
    return r.json()["id"]

# =========================================================
# GET OR CREATE ODOO RECRUITMENT FOLDER
# =========================================================

def get_odoo_folder(models):
    folder = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "search_read",
        [[["name", "=", ODOO_RECRUITMENT_FOLDER], ["type", "=", "folder"]]],
        {"limit": 1}
    )

    if folder:
        return folder[0]["id"]
    return False

# =========================================================
# PROCESS
# =========================================================

def process(payload):

    applicant_id = payload.get("id")
    name = clean(payload.get("display_name") or f"Candidate_{applicant_id}")

    logging.info(f"Processing {name}")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    # =========================
    # ONE DRIVE ROOT
    # =========================

    root_id = get_root()

    # =========================
    # CREATE OR GET FOLDER (IMPORTANT FIX)
    # =========================

    folder_name = name

    search_url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{root_id}/children"
    existing = requests.get(search_url, headers=headers()).json().get("value", [])

    folder_id = None

    for f in existing:
        if f["name"] == folder_name:
            folder_id = f["id"]
            break

    if not folder_id:
        r = requests.post(search_url, json={
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }, headers=headers())

        r.raise_for_status()
        folder_id = r.json()["id"]
   
    # =========================
    # CREATE LINK (ALWAYS WORKS)
    # =========================

    link = requests.post(
        f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink",
        json={"type": "view", "scope": "organization"},
        headers=headers()
    )

    link.raise_for_status()
    weburl = link.json()["link"]["webUrl"]

    # =========================
    # ODOO FOLDER (RECRUITMENT)
    # =========================

    odoo_folder_id = get_odoo_folder(models)

    # =========================
    # CREATE ODOO DOCUMENT
    # =========================

    doc_id = models.execute_kw(
        ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
        "documents.document", "create",
        [{
            "name": f"Candidate Profile - {name}",
            "type": "url",
            "url": weburl,
            "folder_id": odoo_folder_id
        }]
    )

    logging.info(f"Odoo document created {doc_id}")

    # =========================
    # DELETE ATTACHMENTS (SAFE)
    # =========================

    if DELETE_ATTACHMENTS:
        att = payload.get("attachment_ids") or []
        if att:
            models.execute_kw(
                ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
                "ir.attachment", "unlink",
                [att]
            )

# =========================================================
# WEBHOOK
# =========================================================

@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):
    payload = await request.json()
    if payload.get("id"):
        background_tasks.add_task(process, payload)
    return {"ok": True}

@app.get("/")
def health():
    return {"status": "running"}
