import base64
import logging
from fastapi import FastAPI, Request, BackgroundTasks
import xmlrpc.client
import requests

# Configurar logs básicos
logging.basicConfig(level=logging.INFO)

app = FastAPI()

# =====================================================================
# CONFIGURACIÓN DE TUS CREDENCIALES
# =====================================================================
ODOO_URL = "https://yafrel.com"
ODOO_DB = "yafrel-education-platform"
ODOO_USER_ID = 2
ODOO_PASSWORD = "77b329c0854e097fe50cf3660b56830bcd7aecbb"

AZURE_TENANT_ID = "27b934da-5ccf-48d8-8840-f7d7a12209aa"
AZURE_CLIENT_ID = "e82cd831-ab17-4a41-8230-2ae34f22f96f"
AZURE_CLIENT_SECRET = "SEt8Q~j3LYOtatj.z3bXj-Upt-pnk_xkPyXsQbTd"

MICROSOFT_USER_EMAIL = "yafrelservices@yafrel.com"
ODOO_DOCUMENTS_FOLDER_ID = 1  # ID de tu carpeta Reclutamiento

def obtener_token_azure():
    # Corregida la URL de autenticación oficial de Microsoft
    url = f"https://microsoftonline.com{AZURE_TENANT_ID}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': AZURE_CLIENT_ID,
        'client_secret': AZURE_CLIENT_SECRET,
        'scope': 'https://microsoft.com' # Corregido el scope para Microsoft Graph
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()['access_token']

def obtener_o_crear_carpeta_raiz(headers):
    # Corregido el endpoint oficial de Microsoft Graph API
    url = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/root/children"
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
        
        # 1. Asegurar Carpeta Raíz en OneDrive
        raiz_id = obtener_o_crear_carpeta_raiz(headers)

        # 2. Leer datos del postulante desde Odoo usando el ID enviado por el Webhook
        asp = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
            'hr.applicant', 'read', [applicant_id], {'fields': ['name', 'partner_name', 'attachment_ids']})
        
        if not asp:
            logging.error(f"No se encontró el aplicante con ID {applicant_id}")
            return

        nombre_aspirante = asp[0]['partner_name'] or asp[0]['name']
        attachment_ids = asp[0]['attachment_ids']

        # Verificar si ya le creamos un expediente para evitar duplicados
        doc_existe = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
            'documents.document', 'search_count', [[
                ['folder_id', '=', ODOO_DOCUMENTS_FOLDER_ID],
                ['name', '=', f"Expediente - {nombre_aspirante}"]
            ]])

        if doc_existe > 0:
            logging.info(f"El candidato {nombre_aspirante} ya tiene una carpeta vinculada. Saltando.")
            return

        logging.info(f"Procesando nuevo candidato vía Webhook: {nombre_aspirante}")

        # 3. Crear Carpeta del Postulante dentro de "Yafrel Medical Care"
        url_postulante = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{raiz_id}/children"
        postulante_data = {
            "name": nombre_aspirante,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }
        res_postulante = requests.post(url_postulante, json=postulante_data, headers=headers).json()
        postulante_folder_id = res_postulante.get('id')

        # 4. Mover Documentos de Odoo a OneDrive
        if attachment_ids:
            adjuntos = models.execute_kw(ODOO_DB, ODOO_USER_ID, ODOO_PASSWORD,
                'ir.attachment', 'read', [attachment_ids], {'fields': ['name', 'datas']})
                
            for adj in adjuntos:
                file_name = adj['name']
                file_content = base64.b64decode(adj['datas'])
                
                # Subir a OneDrive usando Graph API de flujo binario
                url_upload = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{postulante_folder_id}:/{file_name}:/content"
                headers_file = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/octet-stream'}
                requests.put(url_upload, data=file_content, headers=headers_file)

        # 5. Crear Enlace Compartido en OneDrive
        url_link = f"https://microsoft.com{MICROSOFT_USER_EMAIL}/drive/items/{postulante_folder_id}/createLink"
        link_data = {"type": "view", "scope": "anonymous"}
        res_link = requests.post(url_link, json=link_data, headers=headers).json()
        onedrive_url = res_link.get('link', {}).get('webUrl')

        # 6. Crear el acceso directo tipo Enlace en Odoo Documentos
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
    """
    Recibe la notificación automática de Odoo cuando entra un nuevo postulante
    """
    try:
        payload = await request.json()
        # Odoo envía el ID del registro en el payload primario
        applicant_id = payload.get("id")
        
        if applicant_id:
            # Procesamos en segundo plano para que Odoo no espere y no dé timeout
            background_tasks.add_task(procesar_sincronizacion, int(applicant_id))
            return {"status": "success", "message": "Procesando en segundo plano"}
        
        return {"status": "ignored", "message": "No se detectó ID de postulante"}
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.get("/")
def health_check():
    return {"status": "running", "service": "Yafrel OneDrive Sync"}
