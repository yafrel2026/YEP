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
    url = (
        f"https://login.microsoftonline.com/"
        f"{AZURE_TENANT_ID}/oauth2/v2.0/token"
    )

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
# ONE DRIVE ROOT
# =====================================================================

def obtener_o_crear_carpeta_raiz(headers):

    url = f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/root/children"

    r = requests.get(url, headers=headers)
    r.raise_for_status()

    for item in r.json().get("value", []):
        if item.get("name") == "Yafrel Medical Care":
            return item["id"]

    r = requests.post(url, json={
        "name": "Yafrel Medical Care",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }, headers=headers)

    r.raise_for_status()
    return r.json()["id"]


# =====================================================================
# PROCESAMIENTO PRINCIPAL
# =====================================================================

def procesar_sincronizacion(payload: dict):

    try:
        applicant_id = payload.get("id")

        logging.info(f"Procesando applicant_id={applicant_id}")

        models = xmlrpc.client.ServerProxy(
            f"{ODOO_URL}/xmlrpc/2/object"
        )

        headers = get_headers()
        raiz_id = obtener_o_crear_carpeta_raiz(headers)

        # =========================================================
        # OBTENER DATOS DEL CANDIDATO
        # =========================================================

        aspirante = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "hr.applicant",
            "read",
            [[applicant_id]],
            {"fields": ["name", "partner_name"]}
        )

        if not aspirante:
            logging.error("No se encontró candidato")
            return

        aspirante = aspirante[0]

        nombre = limpiar_nombre(
            aspirante.get("partner_name") or aspirante.get("name")
        )

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
                ["folder_id", "=", ODOO_DOCUMENTS_FOLDER_ID],
                ["name", "=", f"Expediente - {nombre}"]
            ]]
        )

        if doc_existe:
            logging.info("Ya existe expediente")
            return

        # =========================================================
        # CREAR CARPETA EN ONEDRIVE
        # =========================================================

        r = requests.post(
            f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{raiz_id}/children",
            json={
                "name": nombre,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename"
            },
            headers=headers
        )

        r.raise_for_status()
        folder_id = r.json()["id"]

        # =========================================================
        # OBTENER ADJUNTOS (FIX REAL)
        # =========================================================

        attachments = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "ir.attachment",
            "search_read",
            [[
                ["res_model", "=", "hr.applicant"],
                ["res_id", "=", applicant_id]
            ]],
            {"fields": ["name", "datas"]}
        )

        archivos_ok = 0

        for a in attachments:

            if not a.get("datas"):
                continue

            file_name = limpiar_nombre(a["name"])
            content = base64.b64decode(a["datas"])

            url_upload = (
                f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/"
                f"{folder_id}:/{file_name}:/content"
            )

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

        logging.info(f"Archivos subidos: {archivos_ok}")

        # =========================================================
        # CREAR LINK
        # =========================================================

        r = requests.post(
            f"{GRAPH_BASE_URL}/users/{MICROSOFT_USER_EMAIL}/drive/items/{folder_id}/createLink",
            json={"type": "view", "scope": "organization"},
            headers=headers
        )

        r.raise_for_status()

        onedrive_url = r.json()["link"]["webUrl"]

        # =========================================================
        # CREAR DOCUMENT EN ODOO
        # =========================================================

        doc_id = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "create",
            [{
                "name": f"Expediente - {nombre}",
                "type": "url",
                "url": onedrive_url,
                "folder_id": ODOO_DOCUMENTS_FOLDER_ID
            }]
        )

        logging.info(f"Documento creado en Odoo: {doc_id}")

        # =========================================================
        # DELETE ATTACHMENTS (OPCIONAL)
        # =========================================================

        if DELETE_ATTACHMENTS and attachments:

            attachment_ids = [a["id"] for a in attachments]

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "ir.attachment",
                "unlink",
                [attachment_ids]
            )

            logging.info("Adjuntos eliminados")

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
