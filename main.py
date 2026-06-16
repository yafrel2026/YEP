import base64
import logging
import os
import re
import xmlrpc.client

import requests
from fastapi import FastAPI, Request, BackgroundTasks

# =====================================================================
# LOGGING
# =====================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

app = FastAPI()

# =====================================================================
# CONFIGURACIÓN
# =====================================================================

ODOO_URL = "https://yep.yafrel.com"
ODOO_DB = "yafrel-education-platform"

ODOO_USER_ID = int(os.getenv("ODOO_USER_ID", "2"))
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_USER_EMAIL = "yafrelservices@yafrel.com"

ODOO_DOCUMENTS_FOLDER_ID = 1
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

DELETE_ATTACHMENTS = True


# =====================================================================
# VALIDACIÓN
# =====================================================================

required_vars = {
    "ODOO_PASSWORD": ODOO_PASSWORD,
    "AZURE_TENANT_ID": AZURE_TENANT_ID,
    "AZURE_CLIENT_ID": AZURE_CLIENT_ID,
    "AZURE_CLIENT_SECRET": AZURE_CLIENT_SECRET,
}

missing = [k for k, v in required_vars.items() if not v]

if missing:
    raise RuntimeError(f"Faltan variables de entorno: {', '.join(missing)}")


# =====================================================================
# UTILIDADES
# =====================================================================

def limpiar_nombre(nombre: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", nombre).strip()


def obtener_token_azure() -> str:
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"

    data = {
        "grant_type": "client_credentials",
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]


def get_headers():
    token = obtener_token_azure()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


# =====================================================================
# ROOT FOLDER
# =====================================================================

def obtener_o_crear_carpeta_raiz(headers):

    url_list = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"

    r = requests.get(url_list, headers=headers)
    r.raise_for_status()

    for item in r.json().get("value", []):
        if item.get("name") == "Yafrel Medical Care":
            return item["id"]

    url_create = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"

    r = requests.post(url_create, json={
        "name": "Yafrel Medical Care",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }, headers=headers)

    r.raise_for_status()
    return r.json()["id"]


# =====================================================================
# PROCESAMIENTO
# =====================================================================

def procesar_sincronizacion(payload: dict):

    try:
        applicant_id = payload.get("id")

        nombre_aspirante = limpiar_nombre(
            payload.get("display_name", f"Candidato_{applicant_id}")
        )

        attachment_ids = payload.get("attachment_ids", [])

        logging.info(f"Procesando {nombre_aspirante} ({applicant_id})")

        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

        headers = get_headers()
        raiz_id = obtener_o_crear_carpeta_raiz(headers)

        # =========================================================
        # VERIFICAR DUPLICADO EN ODOO
        # =========================================================

        doc_existe = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "search_count",
            [[
                ["name", "=", f"Expediente - {nombre_aspirante}"]
            ]]
        )

        if doc_existe:
            logging.info("Ya existe expediente, skip")
            return

        # =========================================================
        # CARPETA ONEDRIVE
        # =========================================================

        url_folder = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{raiz_id}/children"

        r = requests.post(url_folder, json={
            "name": nombre_aspirante,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }, headers=headers)

        r.raise_for_status()
        folder_id = r.json()["id"]

        # =========================================================
        # UPLOAD
        # =========================================================

        archivos_ok = 0

        if attachment_ids:

            adjuntos = models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "read",
                [attachment_ids],
                {"fields": ["name", "datas"]}
            )

            for a in adjuntos:

                if not a.get("datas"):
                    continue

                file_name = limpiar_nombre(a["name"])
                content = base64.b64decode(a["datas"])

                url_upload = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}:/{file_name}:/content"

                r = requests.put(
                    url_upload,
                    data=content,
                    headers={
                        "Authorization": headers["Authorization"],
                        "Content-Type": "application/octet-stream"
                    }
                )

                r.raise_for_status()
                archivos_ok += 1

        # =========================================================
        # CREATE LINK (BLINDADO)
        # =========================================================

        url_link = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink"

        r = requests.post(url_link, json={
            "type": "view",
            "scope": "organization"
        }, headers=headers)

        r.raise_for_status()

        res = r.json()
        logging.info(f"GRAPH RESPONSE: {res}")

        onedrive_url = (
            res.get("link", {}).get("webUrl")
            or res.get("webUrl")
        )

        if not onedrive_url:
            raise Exception(f"No se pudo obtener URL. Response: {res}")

        # =========================================================
        # ODOO CREATE
        # =========================================================

        doc_id = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "create",
            [{
                "name": f"Expediente - {nombre_aspirante}",
                "type": "url",
                "url": onedrive_url,
                "folder_id": ODOO_DOCUMENTS_FOLDER_ID
            }]
        )

        logging.info(f"ODOO DOC CREATED: {doc_id}")

        # =========================================================
        # DELETE ATTACHMENTS (SOLO SI TODO OK)
        # =========================================================

        if DELETE_ATTACHMENTS and archivos_ok == len(attachment_ids):

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [attachment_ids]
            )

            logging.info("Adjuntos eliminados OK")

    except Exception as e:
        logging.exception(f"ERROR: {str(e)}")


# =====================================================================
# WEBHOOK
# =====================================================================

@app.post("/webhook")
async def webhook(request: Request, background_tasks: BackgroundTasks):

    payload = await request.json()

    if not payload.get("id"):
        return {"status": "ignored"}

    background_tasks.add_task(procesar_sincronizacion, payload)

    return {"status": "ok"}


# =====================================================================
# HEALTH
# =====================================================================

@app.get("/")
def health():
    return {"status": "running"}
