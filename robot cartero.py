# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import os
import json
import re  # Se añade para la reconstrucción estricta de la llave PEM
from google.oauth2 import service_account
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
GOOGLE_SERVICE_ACCOUNT = st.secrets.get("GOOGLE_SERVICE_ACCOUNT")
GEMINI_API_KEY = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")

SCOPES = ['https://www.googleapis.com/auth/drive']
ID_CARPETA_DRIVE = "1kdb74WkObbD94WsKZzRDXfmI4Jlsxgqr"

# --- 3. CONEXIÓN CON GOOGLE DRIVE (CUENTA DE SERVICIO) ---
def autenticar_google():
    if not GOOGLE_SERVICE_ACCOUNT:
        st.warning("⚠️ Configuración de Cuenta de Servicio no detectada. Ejecutando en modo de prueba.")
        return None
    try:
        info_claves = json.loads(GOOGLE_SERVICE_ACCOUNT)
        
        # REPARACIÓN ROBUSTA DE LLAVE PEM:
        # Corrige deformaciones, caracteres de escape '\\n' y saltos incorrectos del portapapeles web
        if "private_key" in info_claves:
            key_original = info_claves["private_key"]
            
            # Limpiamos los encabezados y unificamos todo el cuerpo eliminando rupturas aleatorias
            cuerpo = (key_original
                      .replace("-----BEGIN PRIVATE KEY-----", "")
                      .replace("-----END PRIVATE KEY-----", "")
                      .replace("\\n", "")
                      .replace("\n", "")
                      .replace(" ", "")
                      .strip())
            
            # Forzamos la división exacta de líneas a 64 caracteres de acuerdo a la RFC 1421
            cuerpo_formateado = re.sub(r"(.{64})", r"\1\n", cuerpo)
            
            # Reconstruimos el formato PEM plano ideal para criptografía
            info_claves["private_key"] = f"-----BEGIN PRIVATE KEY-----\n{cuerpo_formateado}\n-----END PRIVATE KEY-----\n"
            
        credenciales = service_account.Credentials.from_service_account_info(
            info_claves, scopes=SCOPES
        )
        return build('drive', 'v3', credentials=credenciales)
    except Exception as e:
        st.error(f"❌ Error al inicializar cuenta de servicio: {e}")
        return None

# Inicialización explícita para evitar NameError en entornos en la nube
drive_service = None
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
if drive_service is not None or not GOOGLE_SERVICE_ACCOUNT:
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
                        st.error("No encontré las plantillas en la carpeta de Drive. Verifica que la carpeta esté compartida con el correo de la Cuenta de Servicio.")
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
