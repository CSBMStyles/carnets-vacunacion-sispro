# Obtener Carnets de Vacunación - MiVacuna SISPRO

Automatiza la descarga del **Carnet Digital de Vacunación** desde:

`https://mivacuna.sispro.gov.co/MiVacuna/Account/LoginCarnetDigitalVac`

Funciona en **macOS** y **Windows**, con interfaz gráfica o por línea de comandos.

---

## Lo que hace

- Lee los empleados desde el Excel de caracterización.
- Permite elegir **hoja**, **fila de encabezados** y **qué columna** tiene cada dato
  (documento, fecha de expedición, nombres, apellidos y perfil).
- Llena el formulario de MiVacuna: tipo de documento `CC`, número, confirmación,
  fecha de expedición y CAPTCHA.
- Resuelve el CAPTCHA con **IA de visión** (cadena de proveedores con respaldo).
- Descarga el PDF como `CEDULA NOMBRES APELLIDOS.pdf` y lo guarda por perfil:
  - `CARNETS DIGITAL ACTUALIZADO/AUXILIAR`
  - `CARNETS DIGITAL ACTUALIZADO/ENFERMERO`
  - `CARNETS DIGITAL ACTUALIZADO/MEDICO`
- Genera `REPORTE_FALLOS_<fecha>.csv` y `.txt` con el motivo de cada fallo.
- Permite trabajar por tandas y continuar donde quedó con
  `estado_proceso_actualizado.json`.

---

## Instalación

### macOS

```bash
cd Codigo/ANTECEDENTES/obtener-carnets-vacunacion-sispro
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp env.example .env
```

Luego abre la interfaz con doble clic en **`iniciar.command`** o:

```bash
python app.py
```

Si macOS bloquea el `chromedriver`:

```bash
xattr -d com.apple.quarantine "$HOME/Library/Application Support/undetected_chromedriver/undetected_chromedriver"
```

### Windows

1. Instala **Python 3.11 o superior** desde <https://www.python.org/downloads/windows/>
   y marca **“Add python.exe to PATH”**.
2. Instala **Google Chrome**.
3. Abre PowerShell en la carpeta del proyecto:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy env.example .env
```

Si PowerShell bloquea la activación:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

4. Abre la interfaz con doble clic en **`iniciar.bat`** o:

```powershell
python app.py
```

> En Linux, `tkinter` se instala aparte: `sudo apt install python3-tk`.

---

## La interfaz

La ventana tiene cuatro pestañas y un registro en la parte inferior.

### 1 · Datos

- **Excel de caracterización**: botón `Examinar…`.
- **Hoja**: se listan las hojas reales del archivo
  (`CONVENIO EJECUCION`, `RETIRO`, …).
- **Encabezados en fila**: por si el Excel los tiene en la fila 13, por ejemplo.
- **Mapeo de columnas**: un desplegable por cada dato. Se detecta automáticamente
  por el nombre del encabezado y tú puedes corregirlo.
- **Vista previa**: muestra las columnas mapeadas primero, para confirmar que la
  cédula quedó en la columna correcta.
- **Validar datos**: cuenta válidos, incompletos y duplicados **sin abrir Chrome**.

### 2 · Salida y navegador

- Carpeta de salida y carpeta temporal de descargas.
- Navegador visible o headless.
- Driver: `auto` (undetected-chromedriver y, si falla, Selenium estándar), `uc` o
  `selenium`.
- Estado de detección de Chrome.

> Si el proyecto está dentro de OneDrive, usa una carpeta temporal local: OneDrive
> puede bloquear el PDF mientras Chrome lo está escribiendo.

### 3 · CAPTCHA / IA

- API keys de OpenCode, Gemini y OpenAI (enmascaradas).
- Cadena de proveedores.
- Modo: `IA`, `IA con respaldo manual`, `Manual`.
- **Probar CAPTCHA**: abre Chrome, captura un CAPTCHA real y muestra qué proveedor
  respondió, el texto y el tiempo.
- **Probar con una imagen…**: prueba la IA con un archivo local, sin navegador.

### 4 · Ejecución

- Perfil (`TODOS`, `AUXILIAR`, `ENFERMERO`, `MEDICO`).
- Cantidad por tanda (`0` = todos).
- Reanudar o reiniciar el estado.
- Intentos, recargas por CAPTCHA roto y espera de descarga.
- Barra de progreso y botón **Detener**.

---

## Configuración

Hay dos archivos, a propósito:

| Archivo | Contenido | ¿Se comparte? |
|---|---|---|
| `config.json` | Rutas, hoja, mapeo de columnas, comportamiento | **No** (local; usa `config.example.json`) |
| `.env` | API keys y ajustes finos de la IA | **No** |

En el repositorio se incluye `config.example.json`; cópialo a `config.json`
o deja que la interfaz lo cree la primera vez que guardes. Si borras el archivo,
la app vuelve a los valores por defecto.

La interfaz guarda todo al iniciar una descarga y al cerrar la ventana.

---

## CAPTCHA: qué necesita para funcionar en Windows

La IA **no corre en tu computador**: se llama por HTTPS. Por eso funciona igual en
macOS y Windows. Lo único que cambia es la **API key**.

### Opción recomendada: Gemini

1. Crea una API key en Google AI Studio.
2. En la pestaña `3 · CAPTCHA / IA`, pégala en **Gemini**.
3. Pulsa **Guardar claves**.
4. Pulsa **Probar CAPTCHA (abre Chrome)**.

Si responde con 4 caracteres y el proveedor es `gemini:...`, quedó listo.

### Respaldo: OpenAI

Si tienes créditos en OpenAI, puedes usarlo como respaldo:

```dotenv
CAPTCHA_PROVIDERS=gemini:gemini-2.5-flash,openai:gpt-4o-mini
```

### Opcional: OpenCode Go / Zen

El proyecto también soporta OpenCode Go/Zen si tienes una key y un modelo de
visión disponible. No es necesario para usarlo: Gemini o el modo manual bastan.

### Si nada funciona: modo manual

En modo `Manual` (o `IA con respaldo manual`), la app muestra la imagen del CAPTCHA
en una ventana y tú escribes los 4 caracteres. Siempre funciona, aunque es más
lento.

### Errores típicos

| Síntoma | Causa probable | Solución |
|---|---|---|
| `401 Unauthorized` / `api_key_expired` | Key vencida, mal copiada o no habilitada para uso directo | Crear otra key o usar Gemini |
| `402` / `429` | Sin fondos o límite de uso | Esperar, usar otro proveedor |
| `No hay OPENCODE_GO_API_KEY` | No hay key y tampoco `auth.json` | Pegar la key en la pestaña 3 |
| `GEMINI_API_KEY vacía` | Falta la key de respaldo | Crearla en AI Studio |
| El CAPTCHA se rechaza 3 veces | El modelo leyó mal los caracteres | Aumentar intentos o cambiar de modelo |

> El script puede leer la key desde `~/.local/share/opencode/auth.json` (macOS/Linux)
> o `%APPDATA%\opencode\auth.json` (Windows), pero **solo si OpenCode está instalado
> en esa máquina**. Para llevar el proyecto a otro equipo, usa una API key en `.env`.

---

## Uso por línea de comandos

La CLI sigue funcionando igual que antes:

```bash
python main.py
```

Modo no interactivo:

```bash
python main.py --perfil AUXILIAR --tanda 10 --modo ia --visible --si
python main.py --perfil TODOS --si
python main.py --reiniciar-estado --si
python main.py --excel "/ruta/CARACTERIZACION.xlsx" --salida "/ruta/salida"
python main.py --hoja "CONVENIO EJECUCION" --fila-encabezado 1 --si
python main.py --dry-run --si
```

`--dry-run` lee y valida el Excel sin abrir el navegador.

---

## Comportamiento del sitio (verificado en vivo)

- IDs del formulario: `tipoIdentificacion`, `numeroIdentificacion`,
  `numeroIdentificacionConfirmacion`, `fechaExpDocumento`, `CaptchaImage`,
  `CaptchaInputText` y el botón `button[name='action:LoginCarnetDigitalVac']`.
- El campo de fecha es un `input` enmascarado (`9999-99-99`) con
  bootstrap-datepicker. Abrir el datepicker puede **cambiar la fecha** escrita
  (`2015-11-24` → `2015-11-29`); por eso se escribe sin abrirlo y se verifica.
- El CAPTCHA a veces sale **roto** (imagen negra con una X roja). Se detecta y se
  recarga la página para pedir otro.
- El enlace **“Cambiar imagen” NO regenera** el CAPTCHA; solo recargar la página.
- Si el CAPTCHA se escribe bien pero el sitio lo rechaza, la página **conserva los
  datos** y emite un CAPTCHA nuevo: el reintento no reescribe nada.
- `send_keys` depende del foco de la ventana (en macOS puede no escribir nada); por
  eso los campos se llenan por JavaScript y se verifican.
- En **headless** el WAF del sitio puede bloquear el navegador (“Web Page Blocked!”)
  porque el User-Agent dice `HeadlessChrome`; el script usa un UA real de Chrome.

---

## Archivos y carpetas

- `app.py`: lanzador de la interfaz.
- `gui.py`: ventana.
- `main.py`: núcleo de descarga y CLI.
- `captcha_solver.py`: proveedores de IA para el CAPTCHA.
- `config.py`: `config.json` y `.env`.
- `sources.py`: lectura del Excel, hojas, mapeo y validación.
- `utils.py`: texto, fechas y nombres de archivo seguros.
- `perfiles.py`: campos y metadatos del proyecto (base para los otros antecedentes).
- `temp_pdf/`: descargas temporales.
- `errors/`: capturas cuando falla una consulta.
- `estado_proceso_actualizado.json`: control de avance.
- `REPORTE_FALLOS_*.csv` / `.txt`: informe de fallos, en la carpeta de salida.

### Motivos del informe

| Motivo | Significado |
|---|---|
| `dato_incompleto` | El registro del Excel no tiene todos los datos. |
| `captcha_incorrecto` | Se agotaron los intentos y el sitio no entregó el PDF. |
| `fecha_no_corresponde` | La fecha no coincide con la del documento. Revisar `Expedicion`. |
| `datos_no_coinciden` | El sitio no encontró coincidencia con documento/nombre. |
| `captcha_ilegible` | El CAPTCHA salió roto o la IA no devolvió texto válido. |
| `error_ia` | Ningún proveedor de IA pudo resolver el CAPTCHA. |
| `formulario_vacio` | El sitio respondió que hay campos obligatorios vacíos. |
| `error_navegador` | Timeout, crash o bloqueo del sitio. |
| `detenido` | Lo detuviste desde la interfaz. |

Los motivos `fecha_no_corresponde` y `datos_no_coinciden` **no se reintentan**: el
problema es el dato, no el CAPTCHA.

---

## Problemas frecuentes

### “This version of ChromeDriver only supports Chrome version …”

`undetected-chromedriver` tiene un driver viejo. El script cae automáticamente a
**Selenium estándar**, que descarga el driver correcto. Si quieres forzarlo:

```bash
DRIVER_BACKEND=selenium
```

o elige `selenium` en la pestaña 2.

### Chrome no aparece

Instala Google Chrome (no basta Edge). La pestaña 2 muestra `Chrome: OK` con la ruta
detectada.

### El PDF no aparece

- Revisa que la carpeta temporal no esté en OneDrive.
- Aumenta la espera de descarga en la pestaña 4.
- Mira la captura en `errors/` y el motivo en el reporte.

### Nombres con acentos o muy largos

El script quita acentos y recorta el nombre a 120 caracteres para no superar el
límite de ruta de Windows. No cambies ese comportamiento si vas a procesar los PDFs
después con otra herramienta.

---

## Diagnósticos

Los scripts internos de diagnóstico no se incluyen en la versión pública.
Para probar el captcha, usa la pestaña **3 · CAPTCHA / IA** de la interfaz.

---

## Reutilizar en los otros antecedentes

`perfiles.py` declara los campos y el tipo de CAPTCHA del proyecto. Para policía,
procuraduría, rethus, etc., se copia la carpeta y se cambia:

- `PROYECTO` (nombre y entidad).
- `CAMPOS` (qué columnas necesita ese trámite).
- El módulo de descarga (`main.py`), conservando la interfaz y `config.py`.

La idea es que la ventana, el guardado de configuración y el manejo de API keys se
reutilicen sin reescribirlos.

---

## Datos y privacidad

Este repositorio contiene únicamente código y documentación. **No incluye**:

- API keys ni archivos `.env`.
- Bases de datos de empleados, cédulas, nombres, correos o teléfonos.
- PDFs, Excel, PPTX, CSV, capturas o reportes generados por las consultas.
- Configuraciones locales (`config.json`) ni estados de proceso.

Para usarlo, proporciona tu propia base de datos autorizada y configura tus
propias credenciales. No subas datos personales al repositorio.

El archivo `config.example.json` muestra una configuración mínima; cópialo a
`config.json` o deja que la interfaz lo cree.

## Licencia

MIT. Consulta el archivo `LICENSE`.

