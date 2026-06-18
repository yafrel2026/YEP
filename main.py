import base64
import logging
import os
import re
import xmlrpc.client

import requests
from fastapi import FastAPI, Request, BackgroundTasks

# --------------------------------------------------
# LOGGING
# --------------------------------------------------

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

ONEDRIVE_ROOT_FOLDER = os.getenv(
    "ONEDRIVE_ROOT_FOLDER",
    "Yafrel Medical Care"
)

ODOO_RECRUITMENT_FOLDER = os.getenv(
    "ODOO_RECRUITMENT_FOLDER",
    "Recruitment"
)

DELETE_ATTACHMENTS = (
    os.getenv("DELETE_ATTACHMENTS", "true").lower() == "true"
)

# --------------------------------------------------

required = [
    ODOO_URL,
    ODOO_DB,
    ODOO_PASSWORD,
    AZURE_TENANT_ID,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET,
    MICROSOFT_DRIVE_USER
]

if not all(required):
    raise RuntimeError(
        "Missing required environment variables"
    )

# --------------------------------------------------


def clean(text):
    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        text
    ).strip()


# --------------------------------------------------


def get_token():

    logging.info("Getting Microsoft token")

    url = (
        f"https://login.microsoftonline.com/"
        f"{AZURE_TENANT_ID}/oauth2/v2.0/token"
    )

    r = requests.post(
        url,
        data={
            "grant_type": "client_credentials",
            "client_id": AZURE_CLIENT_ID,
            "client_secret": AZURE_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default"
        }
    )

    r.raise_for_status()

    logging.info("Microsoft token acquired")

    return r.json()["access_token"]


# --------------------------------------------------


def headers(token):
    return {
        "Authorization": f"Bearer {token}"
    }


# --------------------------------------------------


def get_root_folder(token):

    logging.info(
        f"Looking for root folder '{ONEDRIVE_ROOT_FOLDER}'"
    )

    url = (
        f"{GRAPH}/users/"
        f"{MICROSOFT_DRIVE_USER}"
        f"/drive/root/children"
    )

    r = requests.get(
        url,
        headers=headers(token)
    )

    r.raise_for_status()

    for item in r.json().get("value", []):

        if item["name"] == ONEDRIVE_ROOT_FOLDER:

            logging.info(
                f"Root folder found id={item['id']}"
            )

            return item["id"]

    logging.info(
        "Root folder not found, creating"
    )

    r = requests.post(
        url,
        json={
            "name": ONEDRIVE_ROOT_FOLDER,
            "folder": {}
        },
        headers=headers(token)
    )

    r.raise_for_status()

    folder_id = r.json()["id"]

    logging.info(
        f"Root folder created id={folder_id}"
    )

    return folder_id


# --------------------------------------------------


def get_candidate_folder(
    token,
    root_id,
    candidate_name
):

    logging.info(
        f"Looking for candidate folder '{candidate_name}'"
    )

    url = (
        f"{GRAPH}/users/"
        f"{MICROSOFT_DRIVE_USER}"
        f"/drive/items/{root_id}/children"
    )

    r = requests.get(
        url,
        headers=headers(token)
    )

    r.raise_for_status()

    for item in r.json().get("value", []):

        if item["name"] == candidate_name:

            logging.info(
                f"Candidate folder exists id={item['id']}"
            )

            return item["id"]

    logging.info(
        "Candidate folder not found, creating"
    )

    r = requests.post(
        url,
        json={
            "name": candidate_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        },
        headers=headers(token)
    )

    r.raise_for_status()

    folder_id = r.json()["id"]

    logging.info(
        f"Candidate folder created id={folder_id}"
    )

    return folder_id


# --------------------------------------------------


def process(payload):

    try:

        logging.info(
            "================================================="
        )

        logging.info(
            f"Payload received: {payload}"
        )

        applicant_id = payload.get("id")

        candidate_name = clean(
            payload.get("display_name")
            or f"Candidate_{applicant_id}"
        )

        attachment_ids = payload.get(
            "attachment_ids",
            []
        )

        logging.info(
            f"Candidate: {candidate_name}"
        )

        logging.info(
            f"Attachment IDs: {attachment_ids}"
        )

        token = get_token()

        models = xmlrpc.client.ServerProxy(
            f"{ODOO_URL}/xmlrpc/2/object"
        )

        # --------------------------------------------------
        # ONEDRIVE
        # --------------------------------------------------

        root_id = get_root_folder(token)

        candidate_folder_id = get_candidate_folder(
            token,
            root_id,
            candidate_name
        )

        # --------------------------------------------------
        # ATTACHMENTS
        # --------------------------------------------------

        logging.info(
            "Reading attachments from Odoo"
        )

        attachments = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "ir.attachment",
            "read",
            [attachment_ids],
            {
                "fields": [
                    "id",
                    "name",
                    "datas"
                ]
            }
        )

        logging.info(
            f"Attachments returned by Odoo: {len(attachments)}"
        )

        for att in attachments:

            logging.info(
                f"Attachment "
                f"id={att.get('id')} "
                f"name={att.get('name')} "
                f"datas={'YES' if att.get('datas') else 'NO'}"
            )

        valid_attachments = [
            a for a in attachments
            if a.get("datas")
        ]

        logging.info(
            f"Valid attachments: {len(valid_attachments)}"
        )

        uploaded = 0

        for att in valid_attachments:

            filename = clean(
                att["name"]
            )

            logging.info(
                f"Uploading {filename}"
            )

            content = base64.b64decode(
                att["datas"]
            )

            upload_url = (
                f"{GRAPH}/users/"
                f"{MICROSOFT_DRIVE_USER}"
                f"/drive/items/"
                f"{candidate_folder_id}"
                f":/{filename}:/content"
            )

            r = requests.put(
                upload_url,
                data=content,
                headers={
                    "Authorization":
                        f"Bearer {token}",
                    "Content-Type":
                        "application/octet-stream"
                }
            )

            logging.info(
                f"Upload status={r.status_code}"
            )

            r.raise_for_status()

            uploaded += 1

        logging.info(
            f"Uploaded files: {uploaded}"
        )

        # --------------------------------------------------
        # SHARE LINK
        # --------------------------------------------------

        logging.info(
            "Creating OneDrive link"
        )

        r = requests.post(
            f"{GRAPH}/users/"
            f"{MICROSOFT_DRIVE_USER}"
            f"/drive/items/"
            f"{candidate_folder_id}"
            f"/createLink",
            json={
                "type": "view",
                "scope": "organization"
            },
            headers=headers(token)
        )

        logging.info(
            f"createLink status={r.status_code}"
        )

        r.raise_for_status()

        web_url = r.json()["link"]["webUrl"]

        logging.info(
            f"Link created: {web_url}"
        )

        # --------------------------------------------------
        # DOCUMENTS
        # --------------------------------------------------

        logging.info("Searching Recruitment folder")

        recruitment_folder = models.execute_kw(
        ODOO_DB,
        ODOO_USER_ID,
        ODOO_PASSWORD,
        "documents.document",
        "search_read",
        [[
            ["name", "=", ODOO_RECRUITMENT_FOLDER],
            ["type", "=", "folder"]
        ]],
        {
            "fields": ["id", "name"],
            "limit": 1
        }
        )

        if not recruitment_folder:
        raise Exception("Recruitment folder not found in Odoo")

        recruitment_folder_id = recruitment_folder[0]["id"]

        logging.info(f"Recruitment folder id={recruitment_folder_id}")

        logging.info("Creating URL document")

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
        "folder_id": recruitment_folder_id   # 🔥 ESTE ES EL FIX
        }]
        )

        logging.info(f"Document created id={doc_id}")

        # --------------------------------------------------
        # DELETE
        # --------------------------------------------------

        if (
            DELETE_ATTACHMENTS
            and valid_attachments
            and uploaded == len(valid_attachments)
        ):

            logging.info(
                "Deleting Odoo attachments"
            )

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [attachment_ids]
            )

            logging.info(
                "Attachments deleted"
            )

        logging.info(
            f"SUCCESS: {candidate_name}"
        )

    except Exception as e:

        logging.exception(
            f"PROCESS ERROR: {e}"
        )


# --------------------------------------------------


@app.post("/webhook")
async def webhook(
    request: Request,
    background_tasks: BackgroundTasks
):

    payload = await request.json()

    background_tasks.add_task(
        process,
        payload
    )

    return {"ok": True}


# --------------------------------------------------


@app.get("/")
def health():
    return {
        "status": "running"
    }
