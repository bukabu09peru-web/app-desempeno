"""Equipo por mesa y jefe, fotos de Drive y desempeño de PRODUCTIVIDAD."""
import io
import json
import hashlib
import re
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
ESTRUCTURA_CARPETA_PREDETERMINADA = "1zQmGDrKYU4yl_Uyzy2bk9mOrnWbaAx80"
ACCESOS_PREDETERMINADO = "1RHAWRHEIeRrhJFYzZTzGBGg8GUcR9wxGgwYPrbA7Xs4"
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
SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def clave_mesa(mesa):
    clave = re.sub(r"^mesa\s*", "", normalizar(mesa))
    return str(int(float(clave))) if re.fullmatch(r"\d+(?:\.0+)?", clave) else clave


def interpretar_estructura(filas):
    """Soporta columnas de asesores/promotores y tablas con nombre/rol."""
    alias = {
        "mesa": {"mesa", "mesas"},
        "jefe": {"jefe de credito", "jefe de creditos", "jefe credito", "jefe"},
        "asesor": {"asesor", "asesores"},
        "promotor": {"promotor", "promotores"},
        "nombre": {"nombre", "nombre completo", "colaborador", "colaboradores", "personal", "asesor"},
        "rol": {"rol", "cargo", "puesto"},
    }
    indices, inicio = {}, None
    for i, fila in enumerate(filas[:10]):
        titulos = [normalizar(v) for v in fila]
        hallados = {k: next((j for j, titulo in enumerate(titulos) if titulo in nombres), None)
                    for k, nombres in alias.items()}
        largo = hallados["nombre"] is not None and hallados["rol"] is not None
        ancho = hallados["asesor"] is not None or hallados["promotor"] is not None
        if hallados["mesa"] is not None and hallados["jefe"] is not None and (largo or ancho):
            indices, inicio = hallados, i
            break
    if inicio is None:
        raise ValueError("ESTRUCTURA DE COLABORADORES necesita MESA/MESAS, JEFE DE CREDITO y ASESOR/ASESORES o PROMOTOR/PROMOTORES. También se admite NOMBRE y ROL.")
    def valor(fila, campo):
        j = indices.get(campo)
        return str(fila[j]).strip() if j is not None and j < len(fila) and fila[j] is not None else ""
    resultado, vistos, avisos = [], set(), []
    mesa_actual = jefe_actual = ""
    sin_mesa = sin_jefe = 0
    def agregar(nombre, rol, mesa, jefe):
        for persona in re.split(r"[\n;]+", nombre):
            persona = " ".join(persona.split())
            if not persona:
                continue
            clave = (normalizar(persona), clave_mesa(mesa), normalizar(jefe), rol)
            if clave not in vistos:
                vistos.add(clave)
                resultado.append({"nombre": persona, "mesa": mesa, "jefe": jefe, "rol": rol})
    for fila in filas[inicio+1:]:
        if not any(v is not None and str(v).strip() for v in fila):
            continue
        mesa = valor(fila, "mesa")
        jefe = valor(fila, "jefe")
        if normalizar(mesa) in alias["mesa"] and normalizar(jefe) in alias["jefe"]:
            continue  # encabezado repetido dentro de una tabla larga
        if mesa:
            if clave_mesa(mesa) != clave_mesa(mesa_actual):
                jefe_actual = ""  # nunca heredar el jefe de otra mesa
            mesa_actual = mesa
        if jefe:
            jefe_actual = jefe
        mesa, jefe = mesa_actual, jefe_actual
        if not mesa:
            sin_mesa += 1
            continue
        if jefe:
            agregar(jefe, "JEFE DE CRÉDITO", mesa, jefe)
        if indices["rol"] is not None and indices["nombre"] is not None:
            rol = valor(fila, "rol")
            nr = normalizar(rol)
            rol = ("ASESOR" if "asesor" in nr else "PROMOTOR" if "promotor" in nr
                   else "JEFE DE CRÉDITO" if "jefe" in nr else rol.upper() or "SIN ROL")
            agregar(valor(fila, "nombre"), rol, mesa, jefe)
        else:
            agregar(valor(fila, "asesor"), "ASESOR", mesa, jefe)
            agregar(valor(fila, "promotor"), "PROMOTOR", mesa, jefe)
        if not jefe:
            sin_jefe += 1
    if sin_mesa:
        avisos.append(f"Se omitieron {sin_mesa} filas sin mesa identificable.")
    if sin_jefe:
        avisos.append(f"Hay {sin_jefe} filas de estructura sin jefe asignado; revisa ESTRUCTURA DE COLABORADORES.")
    return resultado, avisos


def escoger_hoja(titulos, preferida):
    exactas = [t for t in titulos if normalizar(t) == normalizar(preferida)]
    if len(exactas) == 1:
        return exactas[0]
    alternativas = [t for t in titulos if normalizar(t) in {"colaboradores", "estructura de colaboradores"}]
    if len(alternativas) == 1:
        return alternativas[0]
    raise ValueError("ACCESOS debe contener la pestaña ESTRUCTURA DE COLABORADORES. Si tiene otro nombre, configúralo como hoja_colaboradores en [app] de Secrets.")


def descargar_excel(drive, archivo):
    limite = 20 * 1024 * 1024
    if int(archivo.get("size", 0)) > limite:
        raise ValueError("ACCESOS supera 20 MB. Usa una copia de menor tamaño o una hoja de Google.")
    buffer = io.BytesIO()
    peticion = drive.files().get_media(fileId=archivo["id"], supportsAllDrives=True)
    lector = MediaIoBaseDownload(buffer, peticion, chunksize=1024 * 1024)
    completo = False
    while not completo:
        _, completo = lector.next_chunk(num_retries=2)
        if buffer.tell() > limite:
            raise ValueError("ACCESOS supera 20 MB.")
    return buffer.getvalue()


def filas_excel(contenido, hoja):
    from openpyxl import load_workbook
    libro = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True, keep_links=False)
    try:
        titulo = escoger_hoja(libro.sheetnames, hoja)
        filas = [["" if v is None else v for v in fila] for fila in libro[titulo].iter_rows(values_only=True)]
        return filas
    finally:
        libro.close()


def localizar_accesos(drive, carpeta_id, archivo_id=""):
    if archivo_id:
        return drive.files().get(fileId=archivo_id, fields="id,name,mimeType,size",
                                 supportsAllDrives=True).execute(num_retries=2)
    carpeta_segura = carpeta_id.replace("\\", "\\\\").replace("'", "\\'")
    candidatos, token = [], None
    while True:
        respuesta = drive.files().list(
            q=f"'{carpeta_segura}' in parents and trashed = false",
            fields="nextPageToken,files(id,name,mimeType,size)", pageSize=1000,
            pageToken=token, spaces="drive", supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute(num_retries=2)
        for f in respuesta.get("files", []):
            nombre = Path(f["name"]).stem if f["name"].lower().endswith((".xlsx", ".xls")) else f["name"]
            if normalizar(nombre) == "accesos" and (f["mimeType"] in {SHEET_MIME, XLSX_MIME}
                    or f["name"].lower().endswith((".xlsx", ".xls"))):
                candidatos.append(f)
        token = respuesta.get("nextPageToken")
        if not token:
            break
    if not candidatos:
        raise ValueError("No se encontró ACCESOS dentro de ESTRUCTURA Y ACCESOS. Comparte la carpeta como Lector con lector-colaboradores y comprueba el nombre del archivo.")
    if len(candidatos) > 1:
        raise ValueError("Hay varios archivos llamados ACCESOS. Configura accesos_id en [app] de Secrets con el ID del archivo que debe usarse.")
    return candidatos[0]


@st.cache_data(ttl=60, max_entries=8, show_spinner=False)
def leer_estructura(carpeta_id, archivo_id, hoja, credenciales_json):
    drive = servicio(credenciales_json)
    archivo = localizar_accesos(drive, carpeta_id, archivo_id)
    if archivo["mimeType"] == SHEET_MIME:
        sheets = servicio(credenciales_json, "sheets", "v4")
        metadata = sheets.spreadsheets().get(spreadsheetId=archivo["id"],
                    fields="sheets(properties(title))").execute(num_retries=2)
        titulo = escoger_hoja([s["properties"]["title"] for s in metadata.get("sheets", [])], hoja)
        filas = sheets.spreadsheets().values().get(spreadsheetId=archivo["id"],
            range="'"+titulo.replace("'", "''")+"'", valueRenderOption="FORMATTED_VALUE"
        ).execute(num_retries=2).get("values", [])
    elif archivo["mimeType"] == XLSX_MIME or archivo["name"].lower().endswith(".xlsx"):
        filas = filas_excel(descargar_excel(drive, archivo), hoja)
    else:
        raise ValueError("Guarda ACCESOS como .xlsx o conviértelo a una hoja de Google. El formato .xls no está admitido.")
    equipo, avisos = interpretar_estructura(filas)
    return equipo, avisos


def error_estructura(error):
    if isinstance(error, HttpError) and error.resp.status in (401, 403, 404):
        return "Comparte ESTRUCTURA Y ACCESOS y su archivo ACCESOS como Lector con la cuenta lector-colaboradores. Comprueba también Google Drive API y Google Sheets API."
    if isinstance(error, ValueError):
        return str(error)
    return "No se pudo leer ESTRUCTURA DE COLABORADORES. Revisa el archivo, su formato y los permisos de lectura."


def construir_equipo(fotos, estructura):
    """Conserva personas sin foto y fotos nuevas aún no asignadas a una mesa."""
    equipo, usadas = [], set()
    nombres_estructura = {normalizar(c["nombre"]) for c in estructura}
    for colaborador in estructura:
        nombre = normalizar(colaborador["nombre"])
        candidatas = [f for f in fotos if normalizar(f["colaborador"]) == nombre]
        if not candidatas:
            aproximadas = [f for f in fotos if coincide_abreviado(nombre, normalizar(f["colaborador"]))]
            if len(aproximadas) == 1:
                foto_nombre = normalizar(aproximadas[0]["colaborador"])
                if len({n for n in nombres_estructura if coincide_abreviado(n, foto_nombre)}) == 1:
                    candidatas = aproximadas
        foto = candidatas[0] if len(candidatas) == 1 else None
        if foto:
            usadas.add(foto["id"])
        c = dict(colaborador, foto=foto)
        clave = [normalizar(c["nombre"]), clave_mesa(c["mesa"]), normalizar(c["jefe"]), c["rol"]]
        c["id"] = "equipo_"+hashlib.sha256(json.dumps(clave).encode()).hexdigest()[:20]
        equipo.append(c)
    for foto in fotos:
        if foto["id"] not in usadas:
            equipo.append({"nombre": foto["colaborador"], "mesa": "", "jefe": "", "rol": "SIN ESTRUCTURA",
                           "foto": foto, "id": "foto_"+foto["id"]})
    return sorted(equipo, key=lambda c: (normalizar(c["mesa"]), normalizar(c["jefe"]), normalizar(c["nombre"]), c["rol"]))


def filtrar_equipo(equipo, mesa="Todas", jefe="Todos", rol="Todos", busqueda=""):
    return [c for c in equipo if (mesa == "Todas" or (mesa == "Sin mesa" and not c["mesa"])
                    or clave_mesa(c["mesa"]) == clave_mesa(mesa))
            and (jefe == "Todos" or (jefe == "Sin jefe asignado" and not c["jefe"])
                 or normalizar(c["jefe"]) == normalizar(jefe))
            and (rol == "Todos" or c["rol"] == rol)
            and normalizar(busqueda) in normalizar(c["nombre"])]


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


def buscar_desempeno(nombre, registros, nombres_fotos, mesa=""):
    # La misma persona puede aparecer en varias mesas. Nunca tomar el desempeño
    # de otra mesa; si la hoja no trae mesa, se permite la coincidencia por nombre.
    if mesa:
        misma_mesa = [r for r in registros if clave_mesa(r["mesa"]) == clave_mesa(mesa)]
        registros = misma_mesa + [r for r in registros if not r["mesa"]]
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
                   else "Sin desempeño asociado a esta persona y mesa. Revisa el nombre en la hoja.")
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
            str(opciones.get("hoja_desempeno", "DESEMPEÑO")),
            str(opciones.get("estructura_carpeta_id", ESTRUCTURA_CARPETA_PREDETERMINADA)),
            str(opciones.get("accesos_id", ACCESOS_PREDETERMINADO)),
            str(opciones.get("hoja_colaboradores", "ESTRUCTURA DE COLABORADORES")))


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


def opciones_unicas(valores, clave=normalizar):
    opciones = {}
    for valor in valores:
        if valor:
            opciones.setdefault(clave(valor), valor)
    return sorted(opciones.values(), key=normalizar)


def filtro(etiqueta, opciones, key):
    # Si el usuario cambia de mesa o se actualiza Drive, la selección anterior
    # puede dejar de existir entre las opciones dependientes.
    if st.session_state.get(key) not in opciones:
        st.session_state[key] = opciones[0]
    return st.selectbox(etiqueta, opciones, key=key)


def main():
    st.set_page_config(page_title="Desempeño · Colaboradores", page_icon="👥", layout="wide")
    st.title("👥 Desempeño del equipo")
    st.caption("Organización por mesa y jefe de crédito · Fotos y desempeño del equipo")
    try:
        (carpeta_id, credenciales, demo, productividad_id, hoja_desempeno,
         estructura_carpeta_id, accesos_id, hoja_colaboradores) = configuracion()
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
            leer_estructura.clear()
        automatico = st.checkbox("Actualizar cada minuto", value=True)
        st.caption("La actualización automática funciona mientras esta página está abierta y activa.")
        st.caption("La estructura se toma de ACCESOS. Las fotos se relacionan por nombre.")
    @st.fragment(run_every=60 if automatico else None)
    def galeria():
        if demo:
            st.info("DEMOSTRACIÓN · Nombres e imágenes ficticios; no se está leyendo Drive.")
            fotos = [{"id": str(i), "colaborador": n} for i, n in enumerate(
                ["Ana Ejemplo", "Carlos Ejemplo", "María Ejemplo"])]
            omitidas = 0
            registros, avisos = interpretar_desempeno([
                ORDEN_ORIGINAL,
                ["MESA 1", "Ana Ejemplo", "5", "12", "3", "4", "1", "3", "S/ 1,000.00", "S/ 700.00", "30%", "S/ 100.00", "S/ 200.00"],
                ["MESA 1", "Carlos Ejemplo", "6", "15", "4", "2", "0", "2", "S/ 1,500.00", "S/ 900.00", "25%", "S/ 150.00", "S/ 250.00"],
                ["MESA 2", "María Ejemplo", "7", "18", "5", "3", "0", "3", "S/ 2,000.00", "S/ 800.00", "40%", "S/ 120.00", "S/ 300.00"],
            ])
            estructura, avisos_estructura = interpretar_estructura([
                ["MESAS", "JEFE DE CREDITO", "ASESORES", "PROMOTORES"],
                ["MESA 1", "Lucía Ejemplo", "Ana Ejemplo", "Luis Ejemplo"],
                ["", "", "Carlos Ejemplo", ""],
                ["MESA 2", "Pedro Ejemplo", "María Ejemplo", ""],
            ])
            fallo_desempeno = False
        else:
            try:
                fotos, omitidas = listar_fotos(carpeta_id, cred_json)
            except Exception as error:
                fotos, omitidas = [], 0
                st.warning("No se pudieron consultar las fotos. " + error_amigable(error))
            try:
                registros, avisos = leer_desempeno(productividad_id, hoja_desempeno, cred_json)
                fallo_desempeno = False
            except Exception as error:
                registros, avisos = [], []
                fallo_desempeno = True
                st.warning("Desempeño no disponible. " + error_productividad(error))
            try:
                estructura, avisos_estructura = leer_estructura(
                    estructura_carpeta_id, accesos_id, hoja_colaboradores, cred_json)
            except Exception as error:
                estructura, avisos_estructura = [], []
                st.warning("Estructura no disponible. " + error_estructura(error))
        for aviso in avisos + avisos_estructura:
            st.info(aviso)
        if omitidas:
            st.warning(f"Hay {omitidas} imágenes con formato no compatible. Usa JPG, PNG, WEBP o GIF.")
        equipo = construir_equipo(fotos, estructura)
        if not equipo:
            st.info("Todavía no hay colaboradores para mostrar. Revisa los avisos, añade personas a ESTRUCTURA DE COLABORADORES o coloca fotos con sus nombres en la carpeta de fotos.")
            return
        nombres = [normalizar(f["colaborador"]) for f in fotos]
        if len(nombres) != len(set(nombres)):
            st.warning("Hay fotos con el mismo nombre. No se asignan a la estructura hasta resolver el duplicado.")
        sin_estructura = sum(c["rol"] == "SIN ESTRUCTURA" for c in equipo)
        if sin_estructura:
            st.caption(f"{sin_estructura} fotos aparecen como SIN ESTRUCTURA. Añade sus nombres a ESTRUCTURA DE COLABORADORES para asignarles mesa, jefe y rol.")
        st.subheader("Filtrar el equipo")
        columnas_filtros = st.columns(3)
        mesas = ["Todas"] + opciones_unicas((c["mesa"] for c in equipo), clave=clave_mesa)
        if any(not c["mesa"] for c in equipo):
            mesas.append("Sin mesa")
        with columnas_filtros[0]:
            mesa = filtro("Mesa", mesas, "filtro_mesa")
        por_mesa = filtrar_equipo(equipo, mesa=mesa)
        jefes = ["Todos"] + opciones_unicas(c["jefe"] for c in por_mesa)
        if any(not c["jefe"] for c in por_mesa):
            jefes.append("Sin jefe asignado")
        with columnas_filtros[1]:
            jefe = filtro("Jefe de crédito", jefes, "filtro_jefe")
        por_jefe = filtrar_equipo(por_mesa, jefe=jefe)
        roles = ["Todos"] + opciones_unicas(c["rol"] for c in por_jefe)
        with columnas_filtros[2]:
            rol = filtro("Rol", roles, "filtro_rol")
        texto = st.text_input("Buscar por nombre", placeholder="Escribe el nombre de un colaborador")
        visibles = filtrar_equipo(equipo, mesa, jefe, rol, texto)
        resumen = st.columns(4)
        for columna, etiqueta, valor in [
            (resumen[0], "Personas", len({normalizar(c["nombre"]) for c in visibles})),
            (resumen[1], "Mesas", len({clave_mesa(c["mesa"]) for c in visibles if c["mesa"]})),
            (resumen[2], "Asesores", len({normalizar(c["nombre"]) for c in visibles if c["rol"] == "ASESOR"})),
            (resumen[3], "Promotores", len({normalizar(c["nombre"]) for c in visibles if c["rol"] == "PROMOTOR"})),
        ]:
            columna.metric(etiqueta, valor)
        if not visibles:
            st.info("No hay colaboradores con estos filtros.")
            return
        with st.expander("Ver estructura del equipo"):
            st.dataframe([{"Mesa": c["mesa"] or "Sin mesa", "Jefe de crédito": c["jefe"] or "Sin jefe asignado",
                           "Colaborador": c["nombre"], "Rol": c["rol"],
                           "Foto": "Disponible" if c["foto"] else "Pendiente"} for c in visibles],
                         hide_index=True, width="stretch")
        nombres_equipo = [c["nombre"] for c in equipo]
        if not fallo_desempeno:
            vinculados = sum(bool(buscar_desempeno(c["nombre"], registros, nombres_equipo, c["mesa"])[0]) for c in visibles)
            st.caption(f"{vinculados} de {len(visibles)} asignaciones visibles tienen desempeño asociado.")
        # Paginación: evita descargar todas las fotos de una carpeta grande.
        paginas = max(1, (len(visibles)+11)//12)
        pagina = filtro("Página", list(range(1, paginas+1)), "pagina")
        st.caption(f"{len(visibles)} asignaciones · Página {pagina} de {paginas}")
        visibles = visibles[(pagina-1)*12:pagina*12]
        columnas = st.columns(3)
        for i, colaborador in enumerate(visibles):
            with columnas[i % 3]:
                with st.container(border=True):
                    archivo = colaborador["foto"]
                    try:
                        if not archivo:
                            contenido = foto_demo(1)
                        elif demo:
                            contenido = foto_demo(int(archivo["id"]))
                        else:
                            if int(archivo.get("size", 0)) > MAX_BYTES:
                                raise ValueError("La foto supera 10 MB. Usa una versión más pequeña.")
                            contenido = descargar_foto(archivo["id"], archivo.get("modifiedTime", ""), cred_json)
                        st.image(contenido, width="stretch")
                    except Exception:
                        st.warning("No se pudo cargar esta foto. Comprueba el formato, tamaño y permisos.")
                    if not archivo:
                        st.caption("Foto pendiente")
                    st.text(colaborador["nombre"])
                    st.caption(colaborador["rol"])
                    st.text("Mesa: " + (colaborador["mesa"] or "Sin mesa"))
                    st.text("Jefe: " + (colaborador["jefe"] or "Sin jefe asignado"))
                    if fallo_desempeno:
                        st.caption("Desempeño no disponible: revisa el aviso de conexión.")
                    else:
                        registros_persona, tipo = buscar_desempeno(colaborador["nombre"], registros, nombres_equipo, colaborador["mesa"])
                        mostrar_indicadores(registros_persona, tipo, colaborador["id"])
        st.caption("ESTRUCTURA DE COLABORADORES define mesa, jefe y rol. Una persona con varias asignaciones aparece en cada una. Las fotos se buscan directamente en la carpeta, sin recorrer subcarpetas.")
    galeria()


if __name__ == "__main__":
    main()
