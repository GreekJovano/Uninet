# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from google import genai
from google.genai import types

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Portal de Automatización de Campo",
    page_icon="⚡",
    layout="centered"
)

st.title("⚡ Automatización de Reportes de Campo")
st.markdown("""
Pega el mensaje de WhatsApp, la Inteligencia Artificial (Gemini) extraerá los datos 
y llenará tus plantillas de Excel guardándolas directamente en tu Google Drive.
""")

# --- 2. CONFIGURACIÓN DE LLAVES Y SEGURIDAD ---
GOOGLE_CLIENT_CONFIG = st.secrets.get("GOOGLE_CLIENT_CONFIG")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

SCOPES = ['https://www.googleapis.com/auth/drive']
REDIRECT_URI = "https://tu-app-en-la-nube.streamlit.app" 
ID_CARPETA_DRIVE = "1kdb74WkObbD94WsKZzRDXfmI4Jlsxgqr"

# --- 3. CONEXIÓN CON GOOGLE DRIVE ---
def autenticar_google():
    if "credentials" not in st.session_state:
        if not GOOGLE_CLIENT_CONFIG:
            st.warning("⚠️ Configuración de Google Drive no detectada. Ejecutando en modo de prueba.")
            return None
        
        flow = Flow.from_client_config(
            json.loads(GOOGLE_CLIENT_CONFIG),
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )
        
        query_params = st.query_params
        if "code" in query_params:
            flow.fetch_token(code=query_params["code"])
            st.session_state.credentials = flow.credentials
            st.rerun()
        else:
            auth_url, _ = flow.authorization_url(prompt='select_account')
            st.markdown(f'<a href="{auth_url}" target="_self" style="display: inline-block; padding: 10px 20px; background-color: #4CAF50; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">🔐 Conectar con Google Drive</a>', unsafe_allow_html=True)
            return None
    
    return build('drive', 'v3', credentials=st.session_state.credentials)

drive_service = autenticar_google()

# --- 4. CEREBRO DE INTELIGENCIA ARTIFICIAL (GEMINI) ---
def analizar_mensaje_con_gemini(texto_whatsapp):
    if not GEMINI_API_KEY:
        st.error("❌ Falta la API Key de Gemini.")
        return None
        
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        
        prompt = f"""
        Analiza el siguiente mensaje de WhatsApp y extrae los campos clave.
        Si algún valor dice 'PENDIENTE' o no está, devuélvelo vacío "".
        
        Mensaje:
        {texto_whatsapp}
        
        Devuelve SOLO un objeto JSON plano con esta estructura:
        {{
            "folio": "...",
            "site_id": "...",
            "cliente": "...",
            "nombre_sitio": "...",
            "domicilio": "...",
            "fecha": "...",
            "hora": "...",
            "actividad": "..."
        }}
        """
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception as e:
        st.error(f"Error con la IA: {e}")
        return None

# --- 5. MOVER ARCHIVOS DESDE Y HACIA GOOGLE DRIVE ---
def descargar_archivo_drive(service, file_id):
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_stream.seek(0)
    return file_stream

def buscar_plantillas_en_drive(service):
    query = f"'{ID_CARPETA_DRIVE}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    return results.get('files', [])

def subir_reporte_a_drive(service, file_stream, file_name):
    file_metadata = {'name': file_name, 'parents': [ID_CARPETA_DRIVE]}
    media = MediaIoBaseUpload(file_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink').execute()

# --- 6. ESCRIBIR EN LOS ENCHUFES/CELDAS DE EXCEL ---
def inyectar_datos_excel(plantilla_stream, datos, tipo_reporte):
    wb = openpyxl.load_workbook(plantilla_stream)
    sheet = wb.active
    nombre = tipo_reporte.lower()
    
    if "uninet" in nombre:
        sheet['B5'] = datos.get('fecha', '')
        sheet['D5'] = datos.get('site_id', '')
        sheet['B6'] = datos.get('cliente', '')
        sheet['B7'] = datos.get('domicilio', '')
        sheet['B9'] = datos.get('actividad', '')
    elif "check" in nombre or "list" in nombre:
        sheet['B4'] = datos.get('cliente', '')
        sheet['B5'] = datos.get('nombre_sitio', '')
        sheet['B6'] = datos.get('domicilio', '')
        sheet['F4'] = datos.get('fecha', '')
    elif "foto" in nombre:
        sheet['A2'] = f"SITIO: {datos.get('nombre_sitio', '')} ({datos.get('site_id', '')})"
        sheet['A3'] = f"FECHA: {datos.get('fecha', '')}"
        
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream

# --- 7. PANTALLA VISUAL (INTERFAZ DE USUARIO) ---
if drive_service or not GOOGLE_CLIENT_CONFIG:
    st.subheader("📥 Entrada de Datos de WhatsApp")
    entrada_texto = st.text_area(
        "Pega aquí el mensaje completo de coordinación:",
        height=180,
        placeholder="FOLIO: PENDIENTE\nSITE ID: 70104404\nCLIENTE: SATFM..."
    )
    
    if st.button("🚀 Procesar y Generar Reportes", type="primary"):
        if not entrada_texto.strip():
            st.warning("Por favor, ingresa un mensaje válido.")
        else:
            with st.spinner("🧠 Leyendo mensaje con Inteligencia Artificial..."):
                datos_json = analizar_mensaje_con_gemini(entrada_texto)
                
            if datos_json:
                st.success("🎯 ¡Datos entendidos!")
                st.json(datos_json)
                
                if drive_service:
                    with st.spinner("📁 Buscando tus archivos en Drive..."):
                        plantillas = buscar_plantillas_en_drive(drive_service)
                        
                    if not plantillas:
                        st.error("No encontré las plantillas en la carpeta de Drive.")
                    else:
                        for p in plantillas:
                            with st.spinner(f"📝 Modificando: {p['name']}..."):
                                stream_plantilla = descargar_archivo_drive(drive_service, p['id'])
                                stream_procesado = inyectar_datos_excel(stream_plantilla, datos_json, p['name'])
                                
                                nuevo_nombre = f"Generado_{datos_json['site_id']}_{p['name']}"
                                resultado = subir_reporte_a_drive(drive_service, stream_procesado, nuevo_nombre)
                                
                                st.markdown(f"✅ Guardado en Drive: **[{nuevo_nombre}]({resultado['webViewLink']})**")
                        st.balloons()
                else:
                    st.info("💡 Modo simulación: Los datos se leyeron pero no se guardó nada en Drive.")