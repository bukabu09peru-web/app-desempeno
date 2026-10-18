"""Fotos de Drive y desempeño por colaborador desde PRODUCTIVIDAD."""
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
PRODUCTIVIDAD_PREDETERMINADA = "1G4lAKm4fxV6ZsPoKyN49Kt5-Ub__hCm6YTvR4GBgIGw"
SCOPE = ["https://www.googleapis.com/auth/drive.readonly",
         "https://www.googleapis.com/auth/spreadsheets.readonly"]
FORMATOS = {"image/jpeg", "image/png", "image/webp", "image/gif"}
MAX_BYTES = 10 * 1024 * 1024
ORDEN_ORIGINAL = ["MESA", "ASESOR", "BOLSAS MEDIO DIA", "BOLSAS CIERRE",
                  "CLIENTES RECUPERADO", "CLIENTES NUEVO", "NO RENOVADOS",
                  "VARIACION CLIENTES", "VARIACIÓN DE CARTERA", "RECUPERACIÓN DE EFECTIVO",
                  "RECUPERACIÓN DE MORA (30%)", "RECUPERACION MORA AGOSTO", "MORA SEPTIEMBRE"]
NOMBRES_ASESOR = {"asesor", "asesores", "colaborador", "colaboradores", "personal",
                  "nombre", "nombre completo"}


def nombre_colaborador(archivo):
    # Juan_Perez.jpg → Juan Perez. Conserva tildes y guiones de nombres compuestos.
    return " ".join(Path(archivo).stem.replace("_", " ").split())


def normalizar(texto):
    texto = str(texto).replace("_", " ").replace("-", " ")
    return " ".join("".join(c for c in unicodedata.normalize("NFD", texto.casefold())
                            if unicodedata.category(c) != "Mn").split())


def servicio(credenciales_json, api="drive", version="v3"):
    credenciales = Credentials.from_service_account_info(
        json.loads(credenciales_json), scopes=SCOPE
    )
    return build(api, version, credentials=credenciales, cache_discovery=False)


def interpretar_desempeno(filas):
    """Detecta el encabezado y conserva los valores tal como se ven en Sheets."""
    if not filas:
        return [], []
    inicio = None
    for numero, fila in enumerate(filas[:10]):
        if any(normalizar(v) in NOMBRES_ASESOR for v in fila):
            inicio = numero
            break
    if inicio is None:
        raise ValueError("No se encontró una columna ASESOR, COLABORADOR, PERSONAL o NOMBRE en las primeras 10 filas de DESEMPEÑO.")
    cabecera = [str(v).strip() for v in filas[inicio]]
    avisos = []
    # El script original explica que A está mal rotulada CLIENTES NUEVO,
    # mientras que B es ASESOR y los indicadores ocupan C:M.
    if (len(cabecera) >= 13 and normalizar(cabecera[0]) == "clientes nuevo"
            and normalizar(cabecera[1]) in NOMBRES_ASESOR):
        cabecera[:13] = ORDEN_ORIGINAL
        avisos.append("Se interpretó A:M con el orden de tu script original porque la primera columna está rotulada CLIENTES NUEVO y contiene la mesa.")
    columna_nombre = next(i for i, v in enumerate(cabecera) if normalizar(v) in NOMBRES_ASESOR)
    columna_mesa = next((i for i, v in enumerate(cabecera) if normalizar(v) in {"mesa", "mesas"}), None)
    resultado = []
    for indice, fila in enumerate(filas[inicio+1:], start=inicio+2):
        nombre = str(fila[columna_nombre]).strip() if columna_nombre < len(fila) else ""
        if not nombre:
            continue
        mesa = str(fila[columna_mesa]).strip() if columna_mesa is not None and columna_mesa < len(fila) else ""
        indicadores, repetidos = {}, {}
        for i, titulo in enumerate(cabecera):
            if i in (columna_nombre, columna_mesa) or not titulo:
                continue
            clave = normalizar(titulo)
            repetidos[clave] = repetidos.get(clave, 0)+1
            etiqueta = titulo if repetidos[clave] == 1 else f"{titulo} ({repetidos[clave]})"
            valor = fila[i] if i < len(fila) else ""
            indicadores[etiqueta] = "—" if valor is None or str(valor).strip() == "" else str(valor)
        resultado.append({"nombre": nombre, "mesa": mesa, "fila": indice, "indicadores": indicadores})
    return resultado, avisos


def coincide_abreviado(a, b):
    return a == b or a.startswith(b + " ") or b.startswith(a + " ")


def buscar_desempeno(nombre, registros, nombres_fotos):
    buscado = normalizar(nombre)
    exactos = [r for r in registros if normalizar(r["nombre"]) == buscado]
    if exactos:
        return exactos, "exacta"
    candidatos = [r for r in registros if coincide_abreviado(buscado, normalizar(r["nombre"]))]
    nombres = {normalizar(r["nombre"]) for r in candidatos}
    if len(nombres) != 1:
        return [], "ambigua" if nombres else "sin datos"
    nombre_origen = next(iter(nombres))
    posibles_fotos = {normalizar(n) for n in nombres_fotos if coincide_abreviado(normalizar(n), nombre_origen)}
    if len(posibles_fotos) != 1:
        return [], "ambigua"
    return candidatos, "abreviada"


@st.cache_data(ttl=60, max_entries=8, show_spinner=False)
def leer_desempeno(archivo_id, hoja, credenciales_json):
    sheets = servicio(credenciales_json, "sheets", "v4")
    metadata = sheets.spreadsheets().get(
        spreadsheetId=archivo_id, fields="sheets(properties(title))"
    ).execute(num_retries=2)
    titulos = [s["properties"]["title"] for s in metadata.get("sheets", [])]
    candidatas = [t for t in titulos if normalizar(t) == normalizar(hoja)]
    if len(candidatas) != 1:
        raise ValueError("No se encontró una pestaña única llamada " + hoja +
                         ". Revisa su nombre; puede configurarse como hoja_desempeno en la sección [app] de Secrets.")
    rango = "'" + candidatas[0].replace("'", "''") + "'"
    respuesta = sheets.spreadsheets().values().get(
        spreadsheetId=archivo_id, range=rango, majorDimension="ROWS",
        valueRenderOption="FORMATTED_VALUE",
    ).execute(num_retries=2)
    return interpretar_desempeno(respuesta.get("values", []))


def error_productividad(error):
    if isinstance(error, HttpError):
        if error.resp.status == 404:
            return "Comparte el archivo PRODUCTIVIDAD como Lector con la cuenta lector-colaboradores y revisa su ID."
        if error.resp.status in (401, 403):
            return "Habilita Google Sheets API en el mismo proyecto y comparte PRODUCTIVIDAD como Lector con la cuenta lector-colaboradores."
        if error.resp.status == 429:
            return "Se alcanzó la cuota de Sheets. Espera un minuto y pulsa Actualizar ahora."
        if error.resp.status == 400:
            return "El archivo debe ser una Hoja de cálculo de Google y tener la pestaña DESEMPEÑO. Si es un Excel, conviértelo a Google Sheets y configura el ID de la copia en productividad_id."
    return str(error) if isinstance(error, ValueError) else "No se pudo consultar PRODUCTIVIDAD. Revisa la conexión y vuelve a actualizar."


def mostrar_indicadores(registros, tipo, archivo_id):
    if not registros:
        st.caption("Nombre ambiguo: haz coincidir el nombre completo de la foto y el de la hoja." if tipo == "ambigua"
                   else "Sin desempeño asociado. Revisa que el nombre coincida con la hoja.")
        return
    if len(registros) > 1:
        st.caption("Hay varios registros para este nombre; selecciona cuál consultar.")
        indice = st.selectbox("Registro de desempeño", list(range(len(registros))), index=None,
                              placeholder="Selecciona una mesa o fila",
                              format_func=lambda i: f"{registros[i]['mesa'] or 'Sin mesa'} · fila {registros[i]['fila']}",
                              key="registro_"+archivo_id)
        if indice is None:
            return
        registro = registros[indice]
    else:
        registro = registros[0]
    if tipo == "abreviada":
        st.caption("Coincidencia por nombre abreviado: " + registro["nombre"])
    if registro["mesa"]:
        st.text("Mesa: " + registro["mesa"])
    indicadores = registro["indicadores"]
    for etiqueta, aliases in [("Bolsas al cierre", {"bolsas cierre"}),
                               ("Clientes nuevos", {"clientes nuevo", "clientes nuevos"}),
                               ("Clientes recuperados", {"clientes recuperado", "clientes recuperados"})]:
        valor = next((v for k, v in indicadores.items() if normalizar(k) in aliases), "—")
        st.metric(etiqueta, valor)
    with st.expander("Ver todos los indicadores"):
        if indicadores:
            st.dataframe([{"Indicador": k, "Valor": v} for k, v in indicadores.items()],
                         hide_index=True, width="stretch")
        else:
            st.info("La fila tiene nombre, pero no tiene columnas de indicadores.")


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
            bool(opciones.get("demo_mode", False)),
            str(opciones.get("productividad_id", PRODUCTIVIDAD_PREDETERMINADA)),
            str(opciones.get("hoja_desempeno", "DESEMPEÑO")))


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
    st.title("👥 Desempeño del equipo")
    st.caption("Fotos de colaboradores y datos de PRODUCTIVIDAD")
    try:
        carpeta_id, credenciales, demo, productividad_id, hoja_desempeno = configuracion()
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
            leer_desempeno.clear()
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
            registros, avisos = interpretar_desempeno([
                ORDEN_ORIGINAL,
                ["MESA DEMO", "Ana Ejemplo", "5", "12", "3", "4", "1", "3", "S/ 1,000.00", "S/ 700.00", "30%", "S/ 100.00", "S/ 200.00"],
                ["MESA DEMO", "Carlos Ejemplo", "6", "15", "4", "2", "0", "2", "S/ 1,500.00", "S/ 900.00", "25%", "S/ 150.00", "S/ 250.00"],
                ["MESA DEMO", "María Ejemplo", "7", "18", "5", "3", "0", "3", "S/ 2,000.00", "S/ 800.00", "40%", "S/ 120.00", "S/ 300.00"],
            ])
            fallo_desempeno = False
        else:
            try:
                fotos, omitidas = listar_fotos(carpeta_id, cred_json)
            except Exception as error:
                st.error(error_amigable(error))
                return
            try:
                registros, avisos = leer_desempeno(productividad_id, hoja_desempeno, cred_json)
                fallo_desempeno = False
            except Exception as error:
                registros, avisos = [], []
                fallo_desempeno = True
                st.warning("Las fotos están disponibles, pero falta el desempeño. " + error_productividad(error))
        for aviso in avisos:
            st.info(aviso)
        st.metric("Fotos de colaboradores", len(fotos))
        if omitidas:
            st.warning(f"Hay {omitidas} imágenes con formato no compatible. Usa JPG, PNG, WEBP o GIF.")
        if not fotos:
            st.info("Todavía no hay fotos compatibles en esta carpeta. Coloca cada foto directamente dentro de ella, por ejemplo: Ana López.jpg.")
            return
        nombres = [normalizar(f["colaborador"]) for f in fotos]
        if len(nombres) != len(set(nombres)):
            st.warning("Hay fotos con el mismo nombre. Cada archivo se muestra como una tarjeta; revisa si hay duplicados.")
        if not fallo_desempeno:
            vinculados = sum(bool(buscar_desempeno(f["colaborador"], registros, [p["colaborador"] for p in fotos])[0]) for f in fotos)
            st.caption(f"{vinculados} de {len(fotos)} fotos tienen coincidencia con {hoja_desempeno}.")
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
                    if fallo_desempeno:
                        st.caption("Desempeño no disponible: revisa el aviso de conexión.")
                    else:
                        registros_persona, tipo = buscar_desempeno(archivo["colaborador"], registros, [p["colaborador"] for p in fotos])
                        mostrar_indicadores(registros_persona, tipo, archivo["id"])
        st.caption("Una foto por colaborador. La galería consulta los archivos ubicados directamente en la carpeta; no recorre subcarpetas ni accesos directos.")
    galeria()


if __name__ == "__main__":
    main()
