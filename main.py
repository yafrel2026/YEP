import base64
import logging
import os
import re
import xmlrpc.client

import requests
from fastapi import FastAPI, Request, BackgroundTasks

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

app = FastAPI()

# --------------------------------------------------
# ENV
# --------------------------------------------------

ODOO_URL = os.getenv("ODOO_URL")
ODOO_DB = os.getenv("ODOO_DB")
ODOO_USER_ID = int(os.getenv("ODOO_USER_ID", "2"))
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_DRIVE_USER = os.getenv("MICROSOFT_DRIVE_USER")

GRAPH = "https://graph.microsoft.com/v1.0"

ONEDRIVE_ROOT_FOLDER = os.getenv("ONEDRIVE_ROOT_FOLDER", "Yafrel Medical Care")
ODOO_RECRUITMENT_FOLDER = os.getenv("ODOO_RECRUITMENT_FOLDER", "Recruitment")

DELETE_ATTACHMENTS = os.getenv("DELETE_ATTACHMENTS", "true").lower() == "true"

# --------------------------------------------------

def clean(text):
    return re.sub(r'[<>:"/\\|?*]', "_", text).strip()

# --------------------------------------------------

def get_recruitment_folder_id(models):

    folder = models.execute_kw(
        ODOO_DB,
        ODOO_USER_ID,
        ODOO_PASSWORD,
        "documents.document",
        "search_read",
        [[
            ["name", "=", ODOO_RECRUITMENT_FOLDER],
            ["type", "=", "folder"]
        ]],
        {"limit": 1}
    )

    if not folder:
        raise Exception("Recruitment folder not found in Odoo Documents")

    return folder[0]["id"]

# --------------------------------------------------

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

# --------------------------------------------------

def headers(token):
    return {"Authorization": f"Bearer {token}"}

# --------------------------------------------------

def process(payload):

    try:

        logging.info("=================================================")
        logging.info(f"Payload received: {payload}")

        applicant_id = payload.get("id")

        candidate_name = clean(
            payload.get("display_name") or f"Candidate_{applicant_id}"
        )

        attachment_ids = payload.get("attachment_ids", [])

        logging.info(f"Candidate: {candidate_name}")
        logging.info(f"Attachment IDs: {attachment_ids}")

        token = get_token()

        models = xmlrpc.client.ServerProxy(
            f"{ODOO_URL}/xmlrpc/2/object"
        )

        # --------------------------------------------------
        # RECRUITMENT FOLDER (FIX CLAVE)
        # --------------------------------------------------

        recruitment_folder_id = get_recruitment_folder_id(models)

        # --------------------------------------------------
        # ATTACHMENTS
        # --------------------------------------------------

        attachments = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "ir.attachment",
            "read",
            [attachment_ids],
            {"fields": ["id", "name", "datas"]}
        )

        valid_attachments = [a for a in attachments if a.get("datas")]

        uploaded = 0

        for att in valid_attachments:

            filename = clean(att["name"])
            content = base64.b64decode(att["datas"])

            logging.info(f"Uploading {filename}")

            upload_url = f"{GRAPH}/users/{MICROSOFT_DRIVE_USER}/drive/items/{candidate_name}:/content"

            r = requests.put(
                upload_url,
                data=content,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream"
                }
            )

            r.raise_for_status()
            uploaded += 1

        # --------------------------------------------------
        # LINK ONEDRIVE
        # --------------------------------------------------

        r = requests.post(
            f"{GRAPH}/users/{MICROSOFT_DRIVE_USER}/drive/items/{candidate_name}/createLink",
            json={"type": "view", "scope": "organization"},
            headers=headers(token)
        )

        r.raise_for_status()
        web_url = r.json()["link"]["webUrl"]

        # --------------------------------------------------
        # 🔥 FIX PRINCIPAL: ODOO DOCUMENT EN RECRUITMENT
        # --------------------------------------------------

        doc_id = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "create",
            [{
                "name": f"Candidate Profile - {candidate_name}",
                "type": "url",
                "url": web_url,
                "folder_id": recruitment_folder_id   # 👈 AQUÍ ESTÁ EL FIX
            }]
        )

        logging.info(f"Document created id={doc_id}")

        # --------------------------------------------------
        # DELETE ATTACHMENTS
        # --------------------------------------------------

        if DELETE_ATTACHMENTS and uploaded == len(valid_attachments):

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [attachment_ids]
            )

            logging.info("Attachments deleted")

        logging.info(f"SUCCESS: {candidate_name}")

    except Exception as e:
        logging.exception(f"PROCESS ERROR: {e}")

# --------------------------------------------------

@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):
    payload = await request.json()
    background_tasks.add_task(process, payload)
    return {"ok": True}

@app.get("/")
def health():
    return {"status": "running"}
