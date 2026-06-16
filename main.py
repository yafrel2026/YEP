import base64
import logging
import os
from fastapi import FastAPI, Request, BackgroundTasks
import xmlrpc.client
import requests

# Configurar logs básicos
logging.basicConfig(level=logging.INFO)

app = FastAPI()

# =====================================================================
# CONFIGURACIÓN SEGURA MEDIANTE VARIABLES DE ENTORNO
# =====================================================================
ODOO_URL = "https://yep.yafrel.com"
ODOO_DB = "yafrel-education-platform"
ODOO_USER_ID = 2
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD")

AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")

MICROSOFT_USER_EMAIL = "yafrelservices@yafrel.com"
ODOO_DOCUMENTS_FOLDER_ID = 1  # ID de tu carpeta Reclutamiento

def obtener_token_azure():
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': AZURE_CLIENT_ID,
        'client_secret': AZURE_CLIENT_SECRET,
        'scope': 'https://microsoft.com'
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()['access_token']

def obtener_o_crear_carpeta_raiz(headers):
    url = f"https://login.microsoft.com{MICROSOFT_USER_EMAIL}/drive/root/children"
    folder_data = {
        "name": "Yafrel Medical Care",
        "folder": {},
        "@microsoft.graph.conflictBehavior": "get"
    }
    res = requests.post(url, json=folder_data, headers=headers).json()
    return res.get('id')

def procesar_sincronizacion(applicant_id: int):
    try:
        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
        token = obtener_token_azure()
        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
        
        raiz_id = obtener_o_crear_carpeta_raiz(headers)

        asp = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
            'hr.applicant', 'read', [applicant_id], {'fields': ['name', 'partner_name', 'attachment_ids']})
        
        if not asp:
            logging.error(f"No se encontró el aplicante con ID {applicant_id}")
            return

        nombre_aspirante = asp['partner_name'] or asp['name']
        attachment_ids = asp['attachment_ids']

        doc_existe = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
            'documents.document', 'search_count', [[
                ['folder_id', '=', ODOO_DOCUMENTS_FOLDER_ID],
                ['name', '=', f"Expediente - {nombre_aspirante}"]
            ]])

        if doc_existe > 0:
            logging.info(f"El candidato {nombre_aspirante} ya tiene una carpeta vinculada. Saltando.")
            return

        logging.info(f"Procesando nuevo candidato vía Webhook: {nombre_aspirante}")

        url_postulante = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{raiz_id}/children"
        postulante_data = {
            "name": nombre_aspirante,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        res_postulante = requests.post(url_postulante, json=postulante_data, headers=headers).json()
        postulante_folder_id = res_postulante.get('id')

        if attachment_ids:
            adjuntos = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
                'ir.attachment', 'read', [attachment_ids], {'fields': ['name', 'datas']})
                
            for adj in adjuntos:
                file_name = adj['name']
                file_content = base64.b64decode(adj['datas'])
                
                url_upload = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{postulante_folder_id}:/{file_name}:/content"
                headers_file = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/octet-stream'}
                requests.put(url_upload, data=file_content, headers=headers_file)

        url_link = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{postulante_folder_id}/createLink"
        link_data = {"type": "view", "scope": "anonymous"}
        res_link = requests.post(url_link, json=link_data, headers=headers).json()
        onedrive_url = res_link.get('link', {}).get('webUrl')

        if onedrive_url:
            document_data = {
                'name': f"Expediente - {nombre_aspirante}",
                'type': 'url',
                'url': onedrive_url,
                'folder_id': ODOO_DOCUMENTS_FOLDER_ID
            }
            models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD, 'documents.document', 'create', [document_data])
            logging.info(f"Carpeta y enlace vinculados con éxito para {nombre_aspirante}")

    except Exception as e:
        logging.error(f"Error durante la sincronización: {str(e)}")

@app.post("/webhook")
async def recibir_odoo_webhook(request: Request, background_tasks: BackgroundTasks):
    try:
        payload = await request.json()
        applicant_id = payload.get("id")
        
        if applicant_id:
            background_tasks.add_task(procesar_sincronizacion, int(applicant_id))
            return {"status": "success", "message": "Procesando en segundo plano"}
        
        return {"status": "ignored", "message": "No se detectó ID de postulante"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/")
def health_check():
    return {"status": "running", "service": "Yafrel OneDrive Sync"}
