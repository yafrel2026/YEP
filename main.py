```python
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

# =====================================================================
# VALIDACIÓN DE VARIABLES DE ENTORNO
# =====================================================================

required_vars = {
    "ODOO_PASSWORD": ODOO_PASSWORD,
    "AZURE_TENANT_ID": AZURE_TENANT_ID,
    "AZURE_CLIENT_ID": AZURE_CLIENT_ID,
    "AZURE_CLIENT_SECRET": AZURE_CLIENT_SECRET,
}

missing = [k for k, v in required_vars.items() if not v]

if missing:
    raise RuntimeError(
        f"Faltan variables de entorno obligatorias: {', '.join(missing)}"
    )


# =====================================================================
# UTILIDADES
# =====================================================================

def limpiar_nombre(nombre: str) -> str:
    """
    Elimina caracteres inválidos para OneDrive.
    """
    return re.sub(r'[<>:"/\\|?*]', "_", nombre).strip()


def obtener_token_azure() -> str:
    """
    Obtiene token OAuth para Microsoft Graph.
    """

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

    response = requests.post(url, data=data)
    response.raise_for_status()

    return response.json()["access_token"]


def obtener_o_crear_carpeta_raiz(headers):
    """
    Crea o recupera la carpeta principal:
    Yafrel Medical Care
    """

    url = (
        f"{GRAPH_BASE_URL}/users/"
        f"{MICROSOFT_USER_EMAIL}/drive/root/children"
    )

    folder_data = {
        "name": "Yafrel Medical Care",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace"
    }

    response = requests.post(
        url,
        json=folder_data,
        headers=headers
    )

    response.raise_for_status()

    data = response.json()

    return data.get("id")


# =====================================================================
# PROCESAMIENTO PRINCIPAL
# =====================================================================

def procesar_sincronizacion(applicant_id: int):

    try:

        logging.info(
            f"Iniciando sincronización para applicant_id={applicant_id}"
        )

        models = xmlrpc.client.ServerProxy(
            f"{ODOO_URL}/xmlrpc/2/object"
        )

        token = obtener_token_azure()

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        raiz_id = obtener_o_crear_carpeta_raiz(headers)

        # =============================================================
        # OBTENER ASPIRANTE
        # =============================================================

        aspirante = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "hr.applicant",
            "read",
            [[applicant_id]],
            {
                "fields": [
                    "name",
                    "partner_name",
                    "attachment_ids"
                ]
            }
        )

        if not aspirante:
            logging.error(
                f"No se encontró el aplicante {applicant_id}"
            )
            return

        aspirante = aspirante[0]

        nombre_aspirante = (
            aspirante.get("partner_name")
            or aspirante.get("name")
            or f"Candidato_{applicant_id}"
        )

        nombre_aspirante = limpiar_nombre(nombre_aspirante)

        attachment_ids = aspirante.get("attachment_ids", [])

        # =============================================================
        # VERIFICAR SI YA EXISTE
        # =============================================================

        doc_existe = models.execute_kw(
            ODOO_DB,
            ODOO_USER_ID,
            ODOO_PASSWORD,
            "documents.document",
            "search_count",
            [[
                [
                    "folder_id",
                    "=",
                    ODOO_DOCUMENTS_FOLDER_ID
                ],
                [
                    "name",
                    "=",
                    f"Expediente - {nombre_aspirante}"
                ]
            ]]
        )

        if doc_existe > 0:
            logging.info(
                f"{nombre_aspirante} ya tiene expediente."
            )
            return

        logging.info(
            f"Procesando candidato: {nombre_aspirante}"
        )

        # =============================================================
        # CREAR CARPETA DEL CANDIDATO
        # =============================================================

        url_postulante = (
            f"{GRAPH_BASE_URL}/users/"
            f"{MICROSOFT_USER_EMAIL}/drive/items/"
            f"{raiz_id}/children"
        )

        postulante_data = {
            "name": nombre_aspirante,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }

        response = requests.post(
            url_postulante,
            json=postulante_data,
            headers=headers
        )

        response.raise_for_status()

        res_postulante = response.json()

        postulante_folder_id = res_postulante.get("id")

        if not postulante_folder_id:
            raise Exception(
                "No fue posible obtener el ID de la carpeta."
            )

        # =============================================================
        # SUBIR ADJUNTOS
        # =============================================================

        if attachment_ids:

            adjuntos = models.execute_kw(
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

            for adjunto in adjuntos:

                file_name = limpiar_nombre(
                    adjunto["name"]
                )

                file_content = base64.b64decode(
                    adjunto["datas"]
                )

                url_upload = (
                    f"{GRAPH_BASE_URL}/users/"
                    f"{MICROSOFT_USER_EMAIL}/drive/items/"
                    f"{postulante_folder_id}:/"
                    f"{file_name}:/content"
                )

                headers_file = {
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/octet-stream"
                }

                upload_response = requests.put(
                    url_upload,
                    data=file_content,
                    headers=headers_file
                )

                upload_response.raise_for_status()

                logging.info(
                    f"Archivo cargado: {file_name}"
                )

        # =============================================================
        # CREAR ENLACE COMPARTIDO
        # =============================================================

        url_link = (
            f"{GRAPH_BASE_URL}/users/"
            f"{MICROSOFT_USER_EMAIL}/drive/items/"
            f"{postulante_folder_id}/createLink"
        )

        link_data = {
            "type": "view",
            "scope": "organization"
        }

        response_link = requests.post(
            url_link,
            json=link_data,
            headers=headers
        )

        response_link.raise_for_status()

        res_link = response_link.json()

        onedrive_url = (
            res_link.get("link", {})
            .get("webUrl")
        )

        # =============================================================
        # CREAR DOCUMENTO EN ODOO
        # =============================================================

        if onedrive_url:

            document_data = {
                "name": f"Expediente - {nombre_aspirante}",
                "type": "url",
                "url": onedrive_url,
                "folder_id": ODOO_DOCUMENTS_FOLDER_ID
            }

            models.execute_kw(
                ODOO_DB,
                ODOO_USER_ID,
                ODOO_PASSWORD,
                "documents.document",
                "create",
                [document_data]
            )

            logging.info(
                f"Expediente creado para "
                f"{nombre_aspirante}"
            )

        else:
            logging.warning(
                "No se recibió URL de OneDrive."
            )

    except requests.exceptions.HTTPError as e:

        logging.error(
            f"Error HTTP Microsoft Graph: {str(e)}"
        )

        if e.response is not None:
            logging.error(e.response.text)

    except Exception as e:

        logging.exception(
            f"Error durante la sincronización: {str(e)}"
        )


# =====================================================================
# WEBHOOK
# =====================================================================

@app.post("/webhook")
async def recibir_odoo_webhook(
    request: Request,
    background_tasks: BackgroundTasks
):

    try:

        payload = await request.json()

        applicant_id = payload.get("id")

        if applicant_id is None:

            return {
                "status": "ignored",
                "message": "No se detectó ID de postulante"
            }

        background_tasks.add_task(
            procesar_sincronizacion,
            int(applicant_id)
        )

        return {
            "status": "success",
            "message": "Procesando en segundo plano"
        }

    except Exception as e:

        logging.exception(
            "Error recibiendo webhook"
        )

        return {
            "status": "error",
            "message": str(e)
        }


# =====================================================================
# HEALTH CHECK
# =====================================================================

@app.get("/")
def health_check():

    return {
        "status": "running",
        "service": "Yafrel OneDrive Sync"
    }
```
