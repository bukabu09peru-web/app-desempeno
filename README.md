# App de desempeño — primera etapa: colaboradores

Esta primera etapa muestra fotos y nombres desde tu carpeta de Google Drive:

https://drive.google.com/drive/folders/1IZXOH9CeQk1XjNq3uvwWl2TkoKMNihcj

Consulta la carpeta cada 60 segundos mientras la página esté abierta y activa.
También tiene un botón Actualizar ahora. No hay que cambiar el código al agregar,
eliminar o renombrar una foto. La aplicación usa permisos de solo lectura.

## 1. Organiza las fotos

- Coloca una foto por colaborador directamente en la carpeta del enlace.
- Nombra cada archivo con el nombre completo: `Ana López.jpg` o `Juan Pérez.png`.
- También se aceptan WEBP y GIF (se muestra el primer cuadro).
- `Ana_López.jpg` se muestra como `Ana López`; se conserva la escritura y las tildes.
- Usa imágenes de menos de 10 MB. Se muestran recortadas en formato cuadrado.
- Las fotos originales no se modifican. Las miniaturas no conservan metadatos EXIF.
- Las subcarpetas y accesos directos no se recorren. Si tu organización es distinta,
  se debe adaptar el código. No uses dos archivos para la misma persona, porque
  cada foto produce una tarjeta.

## 2. Sube el proyecto a GitHub

Usa un repositorio NUEVO, por ejemplo `app-desempeno`; así mantienes tu app de
asistencia separada.

1. Descarga y extrae el ZIP.
2. En GitHub pulsa `+` → `New repository`.
3. Escribe `app-desempeno` y crea el repositorio.
4. Pulsa `Add file` → `Upload files`.
5. Sube el CONTENIDO de la carpeta `desempeno_drive`, no el ZIP.
6. Comprueba que `app.py`, `requirements.txt` y `README.md` estén en la raíz.
7. Sube también `secrets.example.toml`, que solo contiene una plantilla ficticia.
8. La carpeta `.streamlit` es opcional: solo define colores. Windows puede no
   mostrarla si ocultas archivos; la app funciona sin ese tema.
9. Pulsa `Commit changes`.

No subas el archivo JSON de tu cuenta de servicio ni un `secrets.toml` con claves.
Un `.gitignore` ayuda con Git local, pero no sustituye revisar lo que cargas a mano.

## 3. Crea la conexión de Google

Esta conexión es independiente del usuario y contraseña que tenías en tu
dashboard anterior. No necesitas compartir la carpeta públicamente.

1. Abre https://console.cloud.google.com/ con la cuenta que administra tus archivos.
2. En el selector superior de proyectos, crea un proyecto llamado `App desempeño`
   o utiliza un proyecto existente donde tengas permisos.
3. Ve a `APIs y servicios` → `Biblioteca`.
4. Busca **Google Drive API** y pulsa `Habilitar`.
5. Ve a `IAM y administración` → `Cuentas de servicio`.
6. Pulsa `Crear cuenta de servicio`. Ponle `lector-colaboradores`.
7. No necesita un rol de administrador del proyecto. Completa la creación sin
   conceder roles adicionales: el acceso a los archivos se dará compartiendo
   la carpeta en Drive.
8. Abre esa cuenta, entra en `Claves` → `Agregar clave` → `Crear clave nueva`.
9. Elige **JSON** y descarga el archivo. Consérvalo solo en tu equipo.
10. Copia el correo de la cuenta de servicio. Tiene una forma como
    `lector-colaboradores@tu-proyecto.iam.gserviceaccount.com`.
11. Abre tu carpeta en Drive → `Compartir`.
12. Añade ese correo con permiso **Lector**. Comprueba que las fotos hereden el acceso.

Si tu cuenta de trabajo no permite crear claves o compartir con cuentas de
servicio, debes pedir a su administrador que habilite el acceso o usar otra
forma de autenticación. No hace falta hacer pública la carpeta.

## 4. Publica en Streamlit y coloca los secretos

1. En https://share.streamlit.io entra con tu GitHub.
2. Pulsa `Create app` → la opción para desplegar desde GitHub.
3. Selecciona `bukabu09peru-web/app-desempeno`, rama `main`, archivo `app.py`.
   Si elegiste otro repositorio, usa ese nombre.
4. En `Advanced settings`, elige Python **3.12** y abre `Secrets`.
5. Copia allí el contenido de `secrets.example.toml`.
6. Abre el JSON descargado de Google y reemplaza cada valor ficticio por el valor
   del campo con el mismo nombre: `project_id`, `private_key_id`, `private_key`,
   `client_email`, `client_id` y `client_x509_cert_url`.
7. Mantén las comillas y los `\n` de la clave privada. Pega el valor de la clave
   entre comillas dobles en una sola línea, igual que aparece en el JSON.
8. Deja `folder_id` con el ID que ya viene configurado y `demo_mode = false`.
9. Guarda y pulsa `Deploy`.

Si ya publicaste, abre `Manage app` → `Settings` → `Secrets` para configurar esto.
No me envíes la clave privada ni una captura de ella. Si necesitas ayuda, comparte
solo el mensaje de error, sin credenciales.

## 5. Prueba con tus fotos

1. Abre la app. Deben aparecer las fotos con sus nombres.
2. Sube una foto nueva, por ejemplo `Nuevo Colaborador.jpg`, a la carpeta de Drive.
3. Pulsa `Actualizar ahora` o espera el siguiente ciclo de actualización.
4. Renombra o elimina esa foto y vuelve a actualizar para comprobar el cambio.

La comprobación ocurre una vez por minuto mientras la página está abierta y activa;
no es una notificación de fondo y un móvil puede suspender una pestaña en segundo plano.
Los cambios de Google pueden tardar en ser visibles; si no aparece una foto, pulsa
Actualizar ahora y revisa los permisos.

## Modo de demostración opcional

Para ver la interfaz sin conectar ningún dato real, coloca solamente esto en Secrets:

```toml
[app]
demo_mode = true
```

Se mostrarán tres personas ficticias. Antes de conectar Drive, vuelve a la plantilla
completa y usa `demo_mode = false`. Sin credenciales y sin modo demo, la app muestra
instrucciones, nunca fotos reales inventadas.

## Lo que incluye y lo que queda para después

Incluye galería, búsqueda por nombre, paginación de 12 fotos y actualización automática.
Esta etapa no calcula desempeño, no modifica Drive y no tiene login de colaboradores.
Cuando definamos las métricas y sus hojas, se podrán unir a cada colaborador.

Si publicas la app para todo el mundo, sus visitantes podrán ver las fotos que ella
muestra aunque la carpeta de Drive sea privada. Para restringir la app debes limitar
sus visitantes o agregar autenticación antes de compartirla.

## Errores habituales

| Mensaje | Qué revisar |
| --- | --- |
| Falta conectar Google Drive | Coloca la configuración en Secrets, no en GitHub. |
| No se puede acceder a la carpeta o foto | ID correcto; carpeta compartida con `client_email`; permisos heredados. |
| Google denegó el acceso | Habilita Drive API en el proyecto del JSON; revisa permisos. |
| Todavía no hay fotos compatibles | Las fotos deben estar directamente en la carpeta; usa JPG/PNG/WEBP/GIF. |
| No se pudo cargar esta foto | Tamaño inferior a 10 MB; imagen válida; permisos de lectura. |
| Cuota de Google | Espera un minuto; evita pulsar Actualizar repetidamente. |

## Ejecución local opcional

Instala Python 3.12, abre una terminal en esta carpeta y ejecuta:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Para uso local, la configuración real va en `.streamlit/secrets.toml`; ese archivo
está excluido de Git. No renombres `secrets.example.toml` en GitHub con claves reales.

## Validación de esta entrega

Se prueba la interfaz con datos ficticios y la lectura paginada con respuestas de
Google simuladas. La carpeta real necesita tus credenciales y permisos; no se ha
validado su contenido ni una conexión real desde esta entrega.

## Documentación oficial

- Archivos y carpetas: https://developers.google.com/workspace/drive/api/guides/search-files
- Credenciales: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management
- Actualización de pantalla: https://docs.streamlit.io/develop/api-reference/execution-flow/st.fragment
