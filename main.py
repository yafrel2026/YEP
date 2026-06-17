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
# ODOO
# --------------------------------------------------

ODOO_URL = os.getenv("ODOO_URL")
ODOO_DB = os.getenv("ODOO_DB")
ODOO_USER_ID = int(os.getenv("ODOO_USER_ID", "2"))
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

# --------------------------------------------------
# MICROSOFT
# --------------------------------------------------

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_DRIVE_USER = os.getenv("MICROSOFT_DRIVE_USER")

GRAPH = "https://graph.microsoft.com/v1.0"

# --------------------------------------------------
# CONFIG
# --------------------------------------------------

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
    raise RuntimeError("Missing required environment variables")


def clean(text):
    return re.sub(
        r'[<>:"/\\|?*]',
        "_",
        text
    ).strip()


def get_token():

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

    return r.json()["access_token"]


def graph_headers(token):
    return {
        "Authorization": f"Bearer {token}"
    }


def get_root_folder(token):

    url = (
        f"{GRAPH}/users/"
        f"{MICROSOFT_DRIVE_USER}"
        f"/drive/root/children"
    )

    r = requests.get(
        url,
        headers=graph_headers(token)
    )

    r.raise_for_status()

    for item in r.json().get("value", []):
        if item["name"] == ONEDRIVE_ROOT_FOLDER:
            return item["id"]

    r = requests.post(
        url,
        json={
            "name": ONEDRIVE_ROOT_FOLDER,
            "folder": {}
        },
        headers=graph_headers(token)
    )

    r.raise_for_status()

    return r.json()["id"]


def get_candidate_folder(
    token,
    root_id,
    candidate_name
):

    url = (
        f"{GRAPH}/users/"
        f"{MICROSOFT_DRIVE_USER}"
        f"/drive/items/{root_id}/children"
    )

    r = requests.get(
        url,
        headers=graph_headers(token)
    )

    r.raise_for_status()

    for item in r.json().get("value", []):
        if item["name"] == candidate_name:
            return item["id"]

    r = requests.post(
        url,
        json={
            "name": candidate_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        },
        headers=graph_headers(token)
    )

    r.raise_for_status()

    return r.json()["id"]


def get_recruitment_folder(models):

    result = models.execute_kw(
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

    return result[0]["id"] if result else False


def process(payload):

    try:

        applicant_id = payload["id"]

        candidate_name = clean(
            payload.get("display_name")
            or f"Candidate_{applicant_id}"
        )

        attachment_ids = payload.get(
            "attachment_ids",
            []
        )

        logging.info(
            f"Processing {candidate_name}"
        )

        token = get_token()

        models = xmlrpc.client.ServerProxy(
            f"{ODOO_URL}/xmlrpc/2/object"
        )

        # -------------------------
        # ONEDRIVE
        # -------------------------

        root_id = get_root_folder(token)

        candidate_folder_id = get_candidate_folder(
            token,
            root_id,
            candidate_name
        )

        uploaded = 0

        attachments = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "ir.attachment",
            "read",
            [attachment_ids],
            {
                "fields": [
                    "name",
                    "datas"
                ]
            }
        )

        valid_attachments = [
            a for a in attachments
            if a.get("datas")
        ]

        for att in valid_attachments:

            file_name = clean(
                att["name"]
            )

            content = base64.b64decode(
                att["datas"]
            )

            upload_url = (
                f"{GRAPH}/users/"
                f"{MICROSOFT_DRIVE_USER}"
                f"/drive/items/"
                f"{candidate_folder_id}"
                f":/{file_name}:/content"
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

            r.raise_for_status()

            uploaded += 1

        # -------------------------
        # SHARE LINK
        # -------------------------

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
            headers=graph_headers(token)
        )

        r.raise_for_status()

        web_url = r.json()["link"]["webUrl"]

        # -------------------------
        # ODOO DOCUMENTS
        # -------------------------

        recruitment_folder_id = (
            get_recruitment_folder(models)
        )

        if not recruitment_folder_id:
            raise Exception(
                "Recruitment folder not found"
            )

        doc_name = (
            f"Candidate Profile - "
            f"{candidate_name}"
        )

        existing = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "search_count",
            [[
                ["folder_id", "=",
                 recruitment_folder_id],
                ["name", "=",
                 doc_name]
            ]]
        )

        if not existing:

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "documents.document",
                "create",
                [{
                    "name": doc_name,
                    "type": "url",
                    "url": web_url,
                    "folder_id":
                        recruitment_folder_id
                }]
            )

        # -------------------------
        # DELETE ATTACHMENTS
        # -------------------------

        if (
            DELETE_ATTACHMENTS
            and valid_attachments
            and uploaded == len(valid_attachments)
        ):

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [attachment_ids]
            )

            logging.info(
                f"Deleted {len(attachment_ids)} "
                f"attachments from Odoo"
            )

        logging.info(
            f"Completed: {candidate_name}"
        )

    except Exception:
        logging.exception(
            "Applicant processing failed"
        )


@app.post("/webhook")
async def webhook(
    request: Request,
    background_tasks: BackgroundTasks
):

    payload = await request.json()

    if payload.get("id"):
        background_tasks.add_task(
            process,
            payload
        )

    return {"ok": True}


@app.get("/")
def health():
    return {"status": "running"}
