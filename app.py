"""Primera etapa: fotos y nombres de colaboradores desde Google Drive."""
import io
import json
import unicodedata
from pathlib import Path

import streamlit as st
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from PIL import Image, ImageOps

CARPETA_PREDETERMINADA = "1IZXOH9CeQk1XjNq3uvwWl2TkoKMNihcj"
SCOPE = ["https://www.googleapis.com/auth/drive.readonly"]
FORMATOS = {"image/jpeg", "image/png", "image/webp", "image/gif"}
MAX_BYTES = 10 * 1024 * 1024


def nombre_colaborador(archivo):
    # Juan_Perez.jpg → Juan Perez. Conserva tildes y guiones de nombres compuestos.
    return " ".join(Path(archivo).stem.replace("_", " ").split())


def normalizar(texto):
    return "".join(c for c in unicodedata.normalize("NFD", texto.casefold())
                   if unicodedata.category(c) != "Mn")


def servicio(credenciales_json):
    credenciales = Credentials.from_service_account_info(
        json.loads(credenciales_json), scopes=SCOPE
    )
    return build("drive", "v3", credentials=credenciales, cache_discovery=False)


@st.cache_data(ttl=60, max_entries=8, show_spinner=False)
def listar_fotos(carpeta_id, credenciales_json):
    drive = servicio(credenciales_json)
    carpeta = drive.files().get(
        fileId=carpeta_id, fields="name,mimeType", supportsAllDrives=True
    ).execute(num_retries=2)
    if carpeta.get("mimeType") != "application/vnd.google-apps.folder":
        raise ValueError("El ID configurado no corresponde a una carpeta de Google Drive.")
    fotos, omitidas, token = [], 0, None
    carpeta_segura = carpeta_id.replace("\\", "\\\\").replace("'", "\\'")
    while True:
        respuesta = drive.files().list(
            q=f"'{carpeta_segura}' in parents and trashed = false and mimeType contains 'image/'",
            fields="nextPageToken,files(id,name,mimeType,size,modifiedTime)",
            pageSize=1000, pageToken=token, spaces="drive",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute(num_retries=2)
        for archivo in respuesta.get("files", []):
            if archivo["mimeType"] not in FORMATOS:
                omitidas += 1
                continue
            archivo["colaborador"] = nombre_colaborador(archivo["name"])
            fotos.append(archivo)
        token = respuesta.get("nextPageToken")
        if not token:
            break
    return sorted(fotos, key=lambda f: (normalizar(f["colaborador"]), f["id"])), omitidas


@st.cache_data(ttl=3600, max_entries=300, show_spinner=False)
def descargar_foto(archivo_id, modificada, credenciales_json):
    # modificada forma parte de la clave de caché para detectar fotos reemplazadas.
    drive = servicio(credenciales_json)
    peticion = drive.files().get_media(fileId=archivo_id, supportsAllDrives=True)
    salida = io.BytesIO()
    descarga = MediaIoBaseDownload(salida, peticion, chunksize=1024 * 1024)
    terminado = False
    while not terminado:
        _, terminado = descarga.next_chunk(num_retries=2)
        if salida.tell() > MAX_BYTES:
            raise ValueError("La foto supera 10 MB. Sube una versión más pequeña.")
    salida.seek(0)
    with Image.open(salida) as original:
        imagen = ImageOps.exif_transpose(original).convert("RGB")
        imagen = ImageOps.fit(imagen, (450, 450), method=Image.Resampling.LANCZOS)
        miniatura = io.BytesIO()
        imagen.save(miniatura, format="JPEG", quality=85)
        return miniatura.getvalue()


def error_amigable(error):
    if isinstance(error, HttpError):
        if error.resp.status == 404:
            return "No se puede acceder a la carpeta o foto. Revisa su ID y compártela con el correo de la cuenta de servicio."
        if error.resp.status in (401, 403):
            return "Google denegó el acceso. Habilita Google Drive API y revisa los permisos de la cuenta de servicio."
        if error.resp.status == 429:
            return "Se alcanzó la cuota de Google. Espera un minuto y vuelve a intentar."
    if isinstance(error, ValueError):
        return str(error)
    return "No se pudo leer Google Drive. Revisa la conexión y la configuración."


def configuracion():
    try:
        valores = st.secrets.to_dict()
    except FileNotFoundError:
        valores = {}
    opciones = valores.get("app", {})
    return (str(opciones.get("folder_id", CARPETA_PREDETERMINADA)),
            valores.get("gcp_service_account"),
            bool(opciones.get("demo_mode", False)))


def foto_demo(indice):
    colores = ["#087E8B", "#59768B", "#AE647D"]
    img = Image.new("RGB", (450, 450), colores[indice % len(colores)])
    # Silueta sencilla generada por código: datos de prueba, sin fotos reales.
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    draw.ellipse((155, 70, 295, 210), fill="#EAF0F3")
    draw.rounded_rectangle((85, 235, 365, 485), radius=100, fill="#EAF0F3")
    buffer = io.BytesIO()
    img.save(buffer, "PNG")
    return buffer.getvalue()


def main():
    st.set_page_config(page_title="Desempeño · Colaboradores", page_icon="👥", layout="wide")
    st.title("👥 Colaboradores")
    st.caption("Desempeño · Directorio del equipo")
    try:
        carpeta_id, credenciales, demo = configuracion()
    except Exception:
        st.error("La configuración no es válida. Revisa Settings → Secrets.")
        return
    if not demo and not credenciales:
        st.info("Falta conectar Google Drive. Los pasos están en README.md, dentro del proyecto descargado.")
        st.write("**Para conectar la carpeta:**")
        st.write("1. Crea una cuenta de servicio en Google Cloud y habilita Google Drive API.")
        st.write("2. Comparte la carpeta con el correo de esa cuenta, con permiso de Lector.")
        st.write("3. Coloca las credenciales en Settings → Secrets de Streamlit.")
        st.caption("La clave privada se coloca en Secrets; no se sube a GitHub.")
        return
    cred_json = json.dumps(credenciales or {}, sort_keys=True)
    with st.sidebar:
        st.subheader("Directorio")
        if st.button("Actualizar ahora"):
            listar_fotos.clear()
        automatico = st.checkbox("Actualizar cada minuto", value=True)
        st.caption("La actualización automática funciona mientras esta página está abierta y activa.")
        st.caption("Nombre mostrado: nombre de la foto, sin la extensión.")
    @st.fragment(run_every=60 if automatico else None)
    def galeria():
        if demo:
            st.info("DEMOSTRACIÓN · Nombres e imágenes ficticios; no se está leyendo Drive.")
            fotos = [{"id": str(i), "colaborador": n} for i, n in enumerate(
                ["Ana Ejemplo", "Carlos Ejemplo", "María Ejemplo"])]
            omitidas = 0
        else:
            try:
                fotos, omitidas = listar_fotos(carpeta_id, cred_json)
            except Exception as error:
                st.error(error_amigable(error))
                return
        st.metric("Fotos de colaboradores", len(fotos))
        if omitidas:
            st.warning(f"Hay {omitidas} imágenes con formato no compatible. Usa JPG, PNG, WEBP o GIF.")
        if not fotos:
            st.info("Todavía no hay fotos compatibles en esta carpeta. Coloca cada foto directamente dentro de ella, por ejemplo: Ana López.jpg.")
            return
        nombres = [normalizar(f["colaborador"]) for f in fotos]
        if len(nombres) != len(set(nombres)):
            st.warning("Hay fotos con el mismo nombre. Cada archivo se muestra como una tarjeta; revisa si hay duplicados.")
        texto = st.text_input("Buscar por nombre", placeholder="Escribe el nombre de un colaborador")
        visibles = [f for f in fotos if normalizar(texto) in normalizar(f["colaborador"])]
        if not visibles:
            st.info("No hay coincidencias con ese nombre.")
            return
        # Paginación: evita descargar todas las fotos de una carpeta grande.
        paginas = max(1, (len(visibles)+11)//12)
        pagina = st.selectbox("Página", list(range(1, paginas+1)))
        visibles = visibles[(pagina-1)*12:pagina*12]
        columnas = st.columns(3)
        for i, archivo in enumerate(visibles):
            with columnas[i % 3]:
                with st.container(border=True):
                    try:
                        if demo:
                            contenido = foto_demo(int(archivo["id"]))
                        else:
                            if int(archivo.get("size", 0)) > MAX_BYTES:
                                raise ValueError("La foto supera 10 MB. Usa una versión más pequeña.")
                            contenido = descargar_foto(archivo["id"], archivo.get("modifiedTime", ""), cred_json)
                        st.image(contenido, width="stretch")
                    except Exception:
                        st.warning("No se pudo cargar esta foto. Comprueba el formato, tamaño y permisos.")
                    st.text(archivo["colaborador"])
        st.caption("Una foto por colaborador. La galería consulta los archivos ubicados directamente en la carpeta; no recorre subcarpetas ni accesos directos.")
    galeria()


if __name__ == "__main__":
    main()
