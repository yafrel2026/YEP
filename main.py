import base64
import logging
import os
import re
import xmlrpc.client
import requests
from fastapi import FastAPI, Request, BackgroundTasks

# =========================================================
# LOGGING
# =========================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

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

DELETE_ATTACHMENTS = True

ONEDRIVE_ROOT_FOLDER = "Yafrel Medical Care"

ODOO_RECRUITMENT_FOLDER_NAME = "Recruitment"

# =========================================================
# VALIDATION
# =========================================================

if not all([ODOO_PASSWORD, AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
    raise RuntimeError("Missing environment variables")

# =========================================================
# HELPERS
# =========================================================

def clean(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def get_token():
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
    return {"Authorization": f"Bearer {get_token()}"}


# =========================================================
# ONEDRIVE ROOT (YAFREL MEDICAL CARE)
# =========================================================

def get_root_folder():

    url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"

    r = requests.get(url, headers=headers())
    r.raise_for_status()

    for item in r.json().get("value", []):
        if item["name"] == ONEDRIVE_ROOT_FOLDER:
            return item["id"]

    r = requests.post(url, json={
        "name": ONEDRIVE_ROOT_FOLDER,
        "folder": {}
    }, headers=headers())

    r.raise_for_status()
    return r.json()["id"]


# =========================================================
# ODOO RECRUITMENT FOLDER
# =========================================================

def get_odoo_recruitment_folder(models):

    folder = models.execute_kw(
        ODOO_DB,
        ODOO_USER_ID,
        ODOO_PASSWORD,
        "documents.document",
        "search_read",
        [[
            ["name", "=", ODOO_RECRUITMENT_FOLDER_NAME],
            ["type", "=", "folder"]
        ]],
        {"limit": 1}
    )

    if folder:
        return folder[0]["id"]

    return False


# =========================================================
# MAIN PROCESS
# =========================================================

def process(payload):

    applicant_id = payload.get("id")
    name = clean(payload.get("display_name") or f"Candidate_{applicant_id}")

    logging.info(f"Processing {name} ({applicant_id})")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    # =========================
    # ONEDRIVE ROOT
    # =========================

    root_id = get_root_folder()

    # =========================
    # CREATE ONEDRIVE FOLDER
    # =========================

    folder_name = name  # limpio, sin prefijos

    url = f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{root_id}/children"

    r = requests.post(url, json={
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail"
    }, headers=headers())

    if r.status_code not in (200, 201):
        logging.error(f"OneDrive folder error: {r.text}")
        return

    folder_id = r.json()["id"]

    # =========================
    # CREATE SHARE LINK
    # =========================

    link = requests.post(
        f"{GRAPH}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink",
        json={
            "type": "view",
            "scope": "organization"
        },
        headers=headers()
    )

    link.raise_for_status()

    weburl = link.json()["link"]["webUrl"]

    # =========================
    # ODOO RECRUITMENT FOLDER
    # =========================

    odoo_folder_id = get_odoo_recruitment_folder(models)

    # =========================
    # CREATE ODOO DOCUMENT
    # =========================

    doc_id = models.execute_kw(
        ODOO_DB,
        ODOO_USER_ID,
        ODOO_PASSWORD,
        "documents.document",
        "create",
        [{
            "name": f"Candidate Profile - {name}",
            "type": "url",
            "url": weburl,
            "folder_id": odoo_folder_id
        }]
    )

    logging.info(f"Odoo document created: {doc_id}")

    # =========================
    # DELETE ATTACHMENTS (OPTIONAL)
    # =========================

    if DELETE_ATTACHMENTS:
        att = payload.get("attachment_ids") or []

        if att:
            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [att]
            )

            logging.info("Attachments deleted")


# =========================================================
# WEBHOOK
# =========================================================

@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):

    payload = await request.json()

    if payload.get("id"):
        background_tasks.add_task(process, payload)

    return {"ok": True}


# =========================================================
# HEALTH
# =========================================================

@app.get("/")
def health():
    return {"status": "running"}
