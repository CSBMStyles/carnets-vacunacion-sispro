"""
Automatización de descarga del Carnet Digital de Vacunación (MiVacuna - SISPRO).
URL: https://mivacuna.sispro.gov.co/MiVacuna/Account/LoginCarnetDigitalVac

Multiplataforma (Windows / macOS). Ver bloques marcados con [WINDOWS] / [MACOS].
"""

# ---------------------------------------------------------------------------
# IMPORTANTE: en macOS, el Python instalado desde python.org no trae los
# certificados CA, y undetected-chromedriver falla al descargar el chromedriver
# ("CERTIFICATE_VERIFY_FAILED"). Apuntamos SSL_CERT_FILE a certifi ANTES de
# importar undetected_chromedriver. Si certifi no está, no pasa nada.
# ---------------------------------------------------------------------------
import os as _os

try:
    import certifi as _certifi

    _os.environ.setdefault("SSL_CERT_FILE", _certifi.where())
    _os.environ.setdefault("REQUESTS_CA_BUNDLE", _certifi.where())
except Exception:
    pass

import argparse
import base64
import csv
import hashlib
import json
import logging
import os
import platform
import shutil
import time
from datetime import datetime
from pathlib import Path

from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

from captcha_solver import (
    capture_captcha_image,
    is_invalid_captcha,
    normalize_captcha_text,
    solve_captcha_auto,
    solve_captcha_manual,
    valid_captcha_answer,
)

from config import Settings, load_settings
from sources import load_users
from utils import (
    configure_console,
    normalize_text,
    safe_filename,
)

configure_console()

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

VACUNA_URL = "https://mivacuna.sispro.gov.co/MiVacuna/Account/LoginCarnetDigitalVac"

# User-Agent real de Chrome. El WAF de MiVacuna bloquea el navegador cuando
# corre en modo headless (UA "HeadlessChrome"), así que lo sobreescribimos.
CHROME_UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
)
if platform.system() == "Windows":
    CHROME_UA = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/154.0.0.0 Safari/537.36"
    )

PROJECT_DIR = Path(__file__).resolve().parent


class AbortExecution(Exception):
    """Detiene la ejecución completa (modo IA estricto)."""


class SkipUser(Exception):
    """Salta el usuario actual (por ejemplo, en modo manual)."""


# ---------------------------------------------------------------------------
# Configuración activa
# ---------------------------------------------------------------------------

SETTINGS: Settings = load_settings()

EXCEL_PATH = Path(SETTINGS.excel_path)
BASE_OUTPUT_DIR = Path(SETTINGS.output_dir)
STATE_FILE = Path(SETTINGS.state_file)
DOWNLOAD_DIR = Path(SETTINGS.temp_dir)
ERROR_DIR = Path(SETTINGS.error_dir)
CAPTCHA_MAX_ATTEMPTS = int(SETTINGS.captcha_max_attempts)
MAX_IMAGE_RETRIES = int(SETTINGS.max_image_retries)

# Perfiles -> carpeta destino
PROFILE_FOLDER_MAP = {
    "AUXILIAR": "AUXILIAR",
    "ENFERMERO": "ENFERMERO",
    "MEDICO": "MEDICO",
}

LAST_STATS: dict = {}
_STOP_CHECKER = lambda: False


def set_stop_checker(checker=None) -> None:
    """Registra la función que dice si el usuario pidió detener el proceso."""
    global _STOP_CHECKER
    _STOP_CHECKER = checker or (lambda: False)


def stop_requested() -> bool:
    try:
        return bool(_STOP_CHECKER())
    except Exception:
        return False


def apply_settings(settings: Settings) -> Settings:
    """Sincroniza Settings con las variables que usa el resto del módulo.

    La GUI ejecuta un solo trabajo a la vez, así que este único punto de
    sincronización evita pasar la configuración por 20 funciones.
    """
    global SETTINGS, EXCEL_PATH, BASE_OUTPUT_DIR, STATE_FILE, DOWNLOAD_DIR
    global ERROR_DIR, CAPTCHA_MAX_ATTEMPTS, MAX_IMAGE_RETRIES

    SETTINGS = settings
    EXCEL_PATH = Path(settings.excel_path)
    BASE_OUTPUT_DIR = Path(settings.output_dir)
    STATE_FILE = Path(settings.state_file)
    DOWNLOAD_DIR = Path(settings.temp_dir)
    ERROR_DIR = Path(settings.error_dir)
    CAPTCHA_MAX_ATTEMPTS = max(1, int(settings.captcha_max_attempts))
    MAX_IMAGE_RETRIES = max(1, int(settings.max_image_retries))

    if settings.captcha_providers:
        os.environ["CAPTCHA_PROVIDERS"] = settings.captcha_providers

    return SETTINGS


# ---------------------------------------------------------------------------
# Lectura del Excel (delegada a sources.py)
# ---------------------------------------------------------------------------

def load_users_from_excel(excel_path, selected_profiles=None, sheet_name=None,
                          header_row=None, column_map=None):
    """Compatibilidad con la CLI y con versiones anteriores.

    La implementación real está en sources.load_users. Si no se pasan hoja,
    fila de encabezado ni mapeo, se usan los de la configuración activa.
    """
    sheet = SETTINGS.sheet_name if sheet_name is None else sheet_name
    header = SETTINGS.header_row if header_row is None else header_row
    mapping = SETTINGS.column_map if column_map is None else column_map

    users, incomplete, stats = load_users(
        excel_path,
        sheet_name=sheet,
        header_row=header,
        column_map=mapping,
        selected_profiles=selected_profiles,
    )
    globals()["LAST_STATS"] = stats

    logging.info(
        "Columnas detectadas -> Documento: %r | Fecha: %r | Nombre: %r | Apellido: %r | Perfil: %r",
        stats["columnas"].get("documento"), stats["columnas"].get("fecha"),
        stats["columnas"].get("nombre"), stats["columnas"].get("apellido"),
        stats["columnas"].get("perfil"),
    )
    return users, incomplete


# ---------------------------------------------------------------------------
# Estado
# ---------------------------------------------------------------------------

def load_state():
    if not STATE_FILE.exists():
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            data.setdefault("procesados", [])
            data.setdefault("fallidos", [])
            return data
    except Exception:
        return {"procesados": [], "fallidos": [], "ultima_actualizacion": None}


def save_state(state):
    state["ultima_actualizacion"] = datetime.now().isoformat(timespec="seconds")
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(STATE_FILE, "w", encoding="utf-8") as file:
        json.dump(state, file, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Carpetas y navegador
# ---------------------------------------------------------------------------

def configure_download_folders():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ERROR_DIR.mkdir(parents=True, exist_ok=True)
    BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for folder in PROFILE_FOLDER_MAP.values():
        (BASE_OUTPUT_DIR / folder).mkdir(parents=True, exist_ok=True)


def _chrome_options(headless: bool):
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "safebrowsing.enabled": True,
    }
    args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--window-size=1400,1000",
        "--lang=es-CO",
    ]
    if headless:
        args.append("--headless=new")
    return args, prefs


def create_driver(headless=False, backend=None):
    """Crea el driver. Intenta undetected-chromedriver y cae a Selenium normal."""
    backend = (backend or SETTINGS.driver_backend or "auto").strip().lower()
    args, prefs = _chrome_options(headless)

    if backend in ("auto", "uc"):
        try:
            import undetected_chromedriver as uc

            options = uc.ChromeOptions()
            for arg in args:
                options.add_argument(arg)
            options.add_argument(f"--user-agent={CHROME_UA}")
            options.add_experimental_option("prefs", prefs)
            driver = uc.Chrome(options=options)
            logging.info("Driver: undetected-chromedriver")
            return driver
        except Exception as exc:
            if backend == "uc":
                raise
            logging.warning("undetected-chromedriver no disponible (%s). Usando Selenium estándar...", str(exc)[:160])

    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    options = Options()
    for arg in args:
        options.add_argument(arg)
    options.add_argument(f"--user-agent={CHROME_UA}")
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    logging.info("Driver: Selenium estándar")
    return driver


def warmup_page(driver, wait, refresh_times: int = 1):
    """Carga la página y la refresca N veces para reducir detección al inicio."""
    driver.get(VACUNA_URL)
    wait.until(EC.presence_of_element_located((By.ID, "tipoIdentificacion")))
    for _ in range(max(0, refresh_times)):
        time.sleep(1.0)
        driver.refresh()
        wait.until(EC.presence_of_element_located((By.ID, "tipoIdentificacion")))


# ---------------------------------------------------------------------------
# Formulario
# ---------------------------------------------------------------------------

def select_document_type(driver, wait):
    select_elem = wait.until(EC.presence_of_element_located((By.ID, "tipoIdentificacion")))
    select = Select(select_elem)

    # El Excel usa "C.C" -> Cédula de ciudadanía -> value "CC"
    try:
        if any(opt.get_attribute("value") == "CC" for opt in select.options):
            select.select_by_value("CC")
            return
    except Exception:
        pass

    for option in select.options:
        text_norm = normalize_text(option.text)
        if "CEDULA" in text_norm and "CIUDAD" in text_norm:
            select.select_by_visible_text(option.text)
            return

    for option in select.options:
        if option.get_attribute("value"):
            select.select_by_value(option.get_attribute("value"))
            return


JS_SET_VALUE = """
const el = arguments[0], v = arguments[1], esFecha = arguments[2];
let avisos = [];
try {
  if (window.jQuery && window.jQuery.fn.datepicker) {
    // No disparamos focus(): eso ABRE el datepicker y puede sobreescribir la
    // fecha con el día seleccionado (bug real: 2015-11-24 -> 2015-11-29).
    try { window.jQuery(el).datepicker('update', v); } catch (e) { avisos.push('datepicker:' + e.message); }
    try { window.jQuery(el).datepicker('hide'); } catch (e) {}
  }
  if (window.jQuery && window.jQuery(el).inputmask) {
    try { window.jQuery(el).inputmask('setvalue', v); } catch (e) { avisos.push('mask:' + e.message); }
  }
  el.value = v;
  el.dispatchEvent(new Event('input', {bubbles: true}));
  el.dispatchEvent(new Event('change', {bubbles: true}));
  el.dispatchEvent(new Event('blur', {bubbles: true}));
  if (esFecha && el.value !== v) {
    // Último recurso: quitar el datepicker para que no interfiera.
    try { if (window.jQuery && window.jQuery.fn.datepicker) window.jQuery(el).datepicker('remove'); } catch (e) {}
    el.value = v;
    el.dispatchEvent(new Event('input', {bubbles: true}));
    el.dispatchEvent(new Event('change', {bubbles: true}));
    avisos.push('datepicker-removido');
  }
  return {valor: el.value, avisos: avisos};
} catch (e) { return {valor: el.value, avisos: ['error:' + e.message]}; }
"""


def set_input_value(driver, element, value: str, es_fecha: bool = False) -> bool:
    """Escribe un input de forma fiable y verificada.

    - send_keys() depende del foco de la ventana: en macOS, si otra ventana tiene
      el foco, las teclas NO llegan al input. Por eso escribimos por JavaScript.
    - En el campo de fecha no usamos focus(), porque abre el datepicker y puede
      cambiar la fecha escrita.
    """
    value = str(value or "")
    try:
        resultado = driver.execute_script(JS_SET_VALUE, element, value, bool(es_fecha))
        actual = (resultado or {}).get("valor", "")
        if str(actual).strip() == value.strip():
            return True
    except Exception:
        pass

    # Respaldo por teclado (requiere foco)
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        element.clear()
        element.send_keys(value)
        return (element.get_attribute("value") or "").strip() == value.strip()
    except Exception:
        return False


def fill_form(driver, wait, user):
    select_document_type(driver, wait)

    doc_input = wait.until(EC.visibility_of_element_located((By.ID, "numeroIdentificacion")))
    doc_confirm = wait.until(
        EC.visibility_of_element_located((By.ID, "numeroIdentificacionConfirmacion"))
    )
    date_input = wait.until(EC.visibility_of_element_located((By.ID, "fechaExpDocumento")))

    ok_doc = set_input_value(driver, doc_input, user["documento"])
    ok_doc2 = set_input_value(driver, doc_confirm, user["documento"])
    ok_fecha = set_input_value(driver, date_input, str(user["fecha_expedicion"]).strip(), es_fecha=True)

    # Verificación estricta: la fecha debe quedar EXACTAMENTE como el Excel.
    if ok_fecha:
        real = (date_input.get_attribute("value") or "").strip()
        if real != str(user["fecha_expedicion"]).strip():
            logging.warning(
                "La fecha quedó distinta (%r != %r) para %s; reescribiendo...",
                real, user["fecha_expedicion"], user["documento"],
            )
            ok_fecha = set_input_value(
                driver, date_input, str(user["fecha_expedicion"]).strip(), es_fecha=True
            )

    if not (ok_doc and ok_doc2 and ok_fecha):
        logging.warning(
            "No se pudieron escribir todos los campos de %s (doc=%s doc2=%s fecha=%s)",
            user["documento"], ok_doc, ok_doc2, ok_fecha,
        )


def is_form_filled(driver, user) -> bool:
    """Verifica documento, confirmación y fecha EXACTAMENTE como el Excel."""
    try:
        doc_val = (driver.find_element(By.ID, "numeroIdentificacion").get_attribute("value") or "").strip()
        doc2_val = (driver.find_element(By.ID, "numeroIdentificacionConfirmacion").get_attribute("value") or "").strip()
        date_val = (driver.find_element(By.ID, "fechaExpDocumento").get_attribute("value") or "").strip()
        return (
            doc_val == user["documento"].strip()
            and doc2_val == user["documento"].strip()
            and date_val == str(user["fecha_expedicion"]).strip()
        )
    except Exception:
        return False


def ensure_form_filled(driver, wait, user, retries: int = 2) -> bool:
    for _ in range(max(1, retries)):
        if is_form_filled(driver, user):
            return True
        fill_form(driver, wait, user)
        time.sleep(0.35)
    return is_form_filled(driver, user)


JS_SITE_MESSAGE = """
const kw = /(inv[aá]lid|incorrecto|no coincide|no corresponde|no se encontr|obligatorio|no existe|ha expirado|error)/i;
const nodes = document.querySelectorAll(
  '#dvmensaje, #dvmensaje2, .validation-summary-errors, .field-validation-error, .text-danger, .alert, .alert-danger, #divMensaje, span, div, p, label'
);
for (const n of nodes) {
  const t = (n.innerText || '').trim();
  if (t && t.length < 160 && kw.test(t)) {
    return t.replace(/\\s+/g, ' ');
  }
}
return '';
"""


def classify_site_message(mensaje: str) -> str:
    """Traduce el mensaje del sitio a un motivo del informe."""
    texto = normalize_text(mensaje)
    if not texto:
        return ""
    if "FECHA" in texto and ("NO CORRESPONDE" in texto or "EXPEDICION" in texto):
        return "fecha_no_corresponde"
    if "NO COINCIDE" in texto or "NO SE ENCONTR" in texto or "NO EXISTE" in texto:
        return "datos_no_coinciden"
    if "CAPTCHA" in texto or "INVALID" in texto or "INCORRECT" in texto:
        return "captcha_incorrecto"
    if "OBLIGATORIO" in texto or "REQUERID" in texto:
        return "formulario_vacio"
    return ""


# Motivos en los que reintentar NO sirve (el problema es el dato, no el CAPTCHA).
MOTIVOS_SIN_REINTENTO = {"fecha_no_corresponde", "datos_no_coinciden"}


def get_site_message(driver) -> str:
    """Lee el mensaje de error/validación que muestra el sitio.

    Ejemplos reales: "Captcha inválido", "Campo obligatorio",
    "La información ingresada no coincide con nuestros registros".
    """
    try:
        mensaje = driver.execute_script(JS_SITE_MESSAGE)
        if mensaje:
            return str(mensaje)[:300]
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
# CAPTCHA
# ---------------------------------------------------------------------------

JS_FETCH_CAPTCHA = """
const img = document.getElementById('CaptchaImage');
const token = document.getElementById('CaptchaDeText');
const cb = arguments[0];
if (!img) { cb('ERR:sin-imagen'); }
else {
  fetch(img.src, {cache: 'no-store'})
    .then(r => r.blob())
    .then(b => { const fr = new FileReader(); fr.onload = () => cb(fr.result); fr.readAsDataURL(b); })
    .catch(e => cb('ERR:' + e.message));
}
"""


def get_captcha_state(driver):
    """Devuelve (bytes, token) de la imagen actual del CAPTCHA.

    IMPORTANTE: hay que capturar la imagen RENDERIZADA por el navegador
    (element.screenshot). Cuando se pide la misma URL con fetch() el servidor
    devuelve un GIF roto (la X roja). Verificado en vivo.
    """
    token = ""
    try:
        token = driver.execute_script(
            "return document.getElementById('CaptchaDeText') ? document.getElementById('CaptchaDeText').value : '';"
        ) or ""
    except Exception:
        pass

    image_bytes = b""
    try:
        element = driver.find_element(By.ID, "CaptchaImage")
        driver.execute_script(
            "arguments[0].scrollIntoView({block:'center'});", element
        )
        image_bytes = capture_captcha_image(element)
    except Exception:
        image_bytes = b""

    # Respaldo: descargar la imagen por URL (solo si el screenshot falló)
    if not image_bytes:
        try:
            data_url = driver.execute_async_script(JS_FETCH_CAPTCHA)
            if isinstance(data_url, str) and "base64," in data_url:
                image_bytes = base64.b64decode(data_url.split("base64,", 1)[1])
        except Exception:
            image_bytes = b""

    return image_bytes, token


def wait_for_new_captcha(driver, wait, old_token=None, timeout=20):
    """Espera a que la imagen del CAPTCHA esté cargada y (si se pide) sea nueva."""
    def _ready(d):
        try:
            img = d.find_element(By.ID, "CaptchaImage")
            token = d.execute_script("return document.getElementById('CaptchaDeText').value;")
            ready = d.execute_script(
                "return arguments[0].complete && arguments[0].naturalWidth > 0;", img
            )
            if not ready:
                return False
            if old_token and token == old_token:
                return False
            return True
        except Exception:
            return False

    try:
        WebDriverWait(driver, timeout).until(_ready)
    except TimeoutException:
        pass


def reload_login_page(driver, wait):
    """Recarga la página: es la única forma real de obtener un CAPTCHA nuevo.

    (El enlace 'Cambiar imagen' NO regenera la imagen.)
    """
    driver.get(VACUNA_URL)
    wait.until(EC.presence_of_element_located((By.ID, "tipoIdentificacion")))
    wait_for_new_captcha(driver, wait)
    time.sleep(0.3)


def obtain_readable_captcha(driver, wait, user, old_token=None):
    """Obtiene un CAPTCHA legible; recarga si sale negro/vacío. Devuelve (bytes, token)."""
    last_bytes, last_token = b"", ""
    for intento in range(1, MAX_IMAGE_RETRIES + 1):
        if old_token:
            wait_for_new_captcha(driver, wait, old_token=old_token)
        image_bytes, token = get_captcha_state(driver)
        last_bytes, last_token = image_bytes, token

        if image_bytes and not is_invalid_captcha(image_bytes):
            return image_bytes, token

        logging.warning("CAPTCHA ilegible/intento %s (negro o vacío). Recargando página...", intento)
        driver.get(VACUNA_URL)
        wait.until(EC.presence_of_element_located((By.ID, "tipoIdentificacion")))
        fill_form(driver, wait, user)
        time.sleep(0.6)
        old_token = None

    return last_bytes, last_token


# ---------------------------------------------------------------------------
# Descarga
# ---------------------------------------------------------------------------

def wait_for_new_pdf(before_files, timeout=25):
    start = time.time()
    while time.time() - start < timeout:
        current = {f for f in DOWNLOAD_DIR.glob("*.pdf")}
        new_files = list(current - before_files)
        if new_files:
            newest = max(new_files, key=lambda p: p.stat().st_mtime)
            if not (DOWNLOAD_DIR / (newest.name + ".crdownload")).exists():
                return newest
        time.sleep(0.4)
    return None


def build_output_path(user):
    folder = BASE_OUTPUT_DIR / PROFILE_FOLDER_MAP[user["perfil_grupo"]]
    # Se recorta el nombre para no superar MAX_PATH (260) en Windows.
    stem = safe_filename(
        f"{user['documento']} {user['nombre']} {user['apellido']}",
        max_stem=120,
    )
    filename = f"{stem}.pdf"
    target = folder / filename

    if not target.exists():
        return target

    stem, suffix = target.stem, target.suffix
    i = 2
    while True:
        candidate = folder / f"{stem} ({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def save_error_screenshot(driver, user, label="error"):
    try:
        name = f"{label}_{user.get('documento', 'sin_doc')}_{int(time.time())}.png"
        driver.save_screenshot(str(ERROR_DIR / name))
        return str(ERROR_DIR / name)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Flujo por usuario
# ---------------------------------------------------------------------------

def submit_once(driver, wait, user, captcha_mode, rotate=0):
    """Un intento completo: captcha + envío + espera de descarga.

    Devuelve (pdf_path | None, motivo, detalle, captcha_text, token, md5_imagen)
    """
    if not ensure_form_filled(driver, wait, user, retries=2):
        fill_form(driver, wait, user)

    image_bytes, token = obtain_readable_captcha(driver, wait, user)
    fingerprint = hashlib.md5(image_bytes or b"").hexdigest()

    if not image_bytes:
        return None, "captcha_ilegible", "No se pudo capturar la imagen del CAPTCHA", "", token, fingerprint

    if is_invalid_captcha(image_bytes):
        return None, "captcha_ilegible", "La imagen del CAPTCHA sigue negra/vacía", "", token, fingerprint

    # Resolver CAPTCHA
    if captcha_mode == "manual_web":
        captcha_text = solve_captcha_manual(image_bytes, manual_temp_dir=str(ERROR_DIR))
        provider = "manual"
    else:
        try:
            captcha_text, provider = solve_captcha_auto(
                image_bytes, manual_temp_dir=str(ERROR_DIR),
                allow_manual=(captcha_mode == "auto_web"),
                rotate=rotate,
            )
        except Exception as exc:
            return None, "error_ia", f"Ningún proveedor de IA resolvió el CAPTCHA: {str(exc)[:220]}", "", token, fingerprint

    captcha_text = normalize_captcha_text(captcha_text)
    if not valid_captcha_answer(captcha_text):
        return None, "captcha_ilegible", f"Respuesta inválida de {provider}: {captcha_text!r}", captcha_text, token, fingerprint

    # Escribir el CAPTCHA (también por JS: send_keys depende del foco de la ventana)
    captcha_input = wait.until(EC.visibility_of_element_located((By.ID, "CaptchaInputText")))
    if not set_input_value(driver, captcha_input, captcha_text):
        return None, "formulario_vacio", "No se pudo escribir el CAPTCHA", captcha_text, token, fingerprint

    # Verificar TODO antes de enviar: si algo quedó vacío, no se envía.
    if not ensure_form_filled(driver, wait, user, retries=2):
        return None, "formulario_vacio", "El formulario se vació antes de enviar", captcha_text, token, fingerprint

    captcha_actual = (captcha_input.get_attribute("value") or "").strip()
    if captcha_actual != captcha_text:
        return None, "formulario_vacio", (
            f"El campo del CAPTCHA quedó en {captcha_actual!r} en vez de {captcha_text!r}"
        ), captcha_text, token, fingerprint

    before_files = {f for f in DOWNLOAD_DIR.glob("*.pdf")}
    submit_btn = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='action:LoginCarnetDigitalVac']"))
    )
    submit_btn.click()

    pdf_path = wait_for_new_pdf(before_files, timeout=int(SETTINGS.download_timeout or 25))
    if pdf_path:
        return pdf_path, "", f"CAPTCHA resuelto por {provider}", captcha_text, token, fingerprint

    # No descargó: revisar el mensaje del sitio para saber por qué
    time.sleep(0.8)
    site_message = get_site_message(driver)
    motivo = classify_site_message(site_message) or "captcha_incorrecto"
    detalle = site_message or "Sin descarga (posible CAPTCHA incorrecto)"
    return None, motivo, detalle, captcha_text, token, fingerprint


def process_user(driver, wait, user, captcha_mode="auto_web", max_attempts=None):
    """Procesa un usuario con hasta N intentos.

    Comportamiento verificado en el sitio: cuando el CAPTCHA es rechazado, la
    página CONSERVA los datos ingresados y emite un CAPTCHA nuevo. Por eso, tras
    un fallo no recargamos ni volvemos a escribir: solo leemos el CAPTCHA nuevo
    y reenviamos. Solo se recarga si el formulario quedó vacío o si el CAPTCHA
    nuevo también salió roto.
    """
    max_attempts = max_attempts or CAPTCHA_MAX_ATTEMPTS
    intentos = []
    last_reason, last_detail = "sin_descarga", ""
    last_token, last_fingerprint = None, None

    for attempt in range(1, max_attempts + 1):
        if stop_requested():
            last_reason, last_detail = "detenido", "El usuario detuvo el proceso."
            break
        try:
            if attempt == 1:
                reload_login_page(driver, wait)
                fill_form(driver, wait, user)
            else:
                # Tras un fallo el sitio puede CONSERVAR los datos y emitir un
                # CAPTCHA nuevo. Solo recargamos si los datos se vaciaron, si el
                # CAPTCHA no cambió (sería el mismo) o si salió roto.
                imagen_actual, token_actual = get_captcha_state(driver)
                mismo_captcha = (
                    token_actual == last_token
                    and hashlib.md5(imagen_actual or b"").hexdigest() == last_fingerprint
                )
                if not is_form_filled(driver, user) or mismo_captcha or is_invalid_captcha(imagen_actual):
                    reload_login_page(driver, wait)
                    fill_form(driver, wait, user)

            pdf_path, reason, detail, captcha_text, token, fingerprint = submit_once(
                driver, wait, user, captcha_mode, rotate=attempt - 1
            )
            last_token, last_fingerprint = token, fingerprint
            intentos.append({
                "intento": attempt,
                "captcha": captcha_text,
                "resultado": "OK" if pdf_path else reason,
                "detalle": detail,
            })

            if pdf_path:
                output_path = build_output_path(user)
                shutil.move(str(pdf_path), str(output_path))
                logging.info(
                    "[OK] Guardado: %s | Perfil: %s | Doc: %s (intento %s)",
                    output_path.name, user["perfil_grupo"], user["documento"], attempt,
                )
                return True, intentos, "", "", ""

            last_reason, last_detail = reason, detail
            logging.warning(
                "[WARN] Intento %s/%s falló para %s: %s", attempt, max_attempts, user["documento"], detail
            )

            if reason in MOTIVOS_SIN_REINTENTO:
                logging.warning(
                    "[SKIP] No se reintenta %s: el problema es el dato, no el CAPTCHA (%s).",
                    user["documento"], reason,
                )
                break

        except (TimeoutException, WebDriverException) as exc:
            last_reason, last_detail = "error_navegador", f"{type(exc).__name__}: {str(exc)[:200]}"
            intentos.append({"intento": attempt, "captcha": "", "resultado": last_reason, "detalle": last_detail})
            logging.error("Error de navegador con %s: %s", user["documento"], last_detail)
            try:
                reload_login_page(driver, wait)
            except Exception:
                pass
        except SkipUser:
            raise
        except Exception as exc:
            last_reason, last_detail = "error_inesperado", f"{type(exc).__name__}: {str(exc)[:200]}"
            intentos.append({"intento": attempt, "captcha": "", "resultado": last_reason, "detalle": last_detail})
            logging.error("Error inesperado con %s: %s", user["documento"], last_detail)

    if last_reason == "detenido":
        return False, intentos, last_reason, last_detail, ""

    screenshot = save_error_screenshot(driver, user, label=last_reason)
    return False, intentos, last_reason, last_detail, screenshot


# ---------------------------------------------------------------------------
# Informe de fallos
# ---------------------------------------------------------------------------

def write_failure_report(fallos, incomplete_rows, timestamp):
    if not fallos and not incomplete_rows:
        return None, None

    csv_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.csv"
    txt_path = BASE_OUTPUT_DIR / f"REPORTE_FALLOS_{timestamp}.txt"

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "documento", "nombre", "perfil", "fila_excel", "motivo", "detalle",
            "intentos", "captcha(s) intentado(s)", "fecha", "captura",
        ])
        for fila in incomplete_rows:
            writer.writerow([
                fila.get("documento", ""), fila.get("nombre", ""), "", fila.get("fila", ""),
                fila.get("motivo", "dato_incompleto"), fila.get("detalle", ""), 0, "", timestamp, "",
            ])
        for fallo in fallos:
            captchas = " | ".join(
                f"[{i['intento']}] {i['captcha'] or '-'} ({i['resultado']})" for i in fallo["intentos"]
            )
            writer.writerow([
                fallo["documento"], fallo["nombre"], fallo["perfil"], fallo.get("fila_excel", ""),
                fallo["motivo"], fallo["detalle"], len(fallo["intentos"]), captchas, timestamp,
                fallo.get("captura", ""),
            ])

    from collections import Counter

    contador = Counter(f["motivo"] for f in fallos)
    contador.update(f.get("motivo", "dato_incompleto") for f in incomplete_rows)

    lineas = [
        "REPORTE DE FALLOS - DESCARGA CARNET DIGITAL VACUNACIÓN (MiVacuna)",
        f"Generado: {timestamp}",
        f"Excel: {EXCEL_PATH}",
        f"Carpeta de salida: {BASE_OUTPUT_DIR}",
        "",
        f"Total de fallos: {len(fallos) + len(incomplete_rows)}",
        "",
        "RESUMEN POR MOTIVO",
    ]
    descripcion_motivos = {
        "dato_incompleto": "El registro del Excel no tiene todos los datos requeridos.",
        "captcha_incorrecto": "Se agotaron los intentos y el sitio no entregó el PDF (CAPTCHA rechazado).",
        "fecha_no_corresponde": "El sitio dice que la fecha de expedición no corresponde con la del documento. Revisar el dato en el Excel (columna 'Expedicion').",
        "datos_no_coinciden": "El sitio no encontró coincidencia con los datos (documento/nombre). Verificar el registro.",
        "captcha_ilegible": "La imagen del CAPTCHA no se pudo leer (salió negra/vacía o la IA no devolvió texto válido).",
        "error_ia": "Ningún proveedor de IA pudo resolver el CAPTCHA (error de API/cuota).",
        "formulario_vacio": "El sitio respondió que hay campos obligatorios vacíos.",
        "error_navegador": "Fallo del navegador (timeout, crash o bloqueo del sitio).",
        "error_inesperado": "Error inesperado durante el proceso.",
        "sin_descarga": "El sitio no entregó el PDF y no se leyó un mensaje específico.",
    }
    for motivo, cantidad in contador.most_common():
        lineas.append(f"  - {motivo}: {cantidad}  ({descripcion_motivos.get(motivo, 'Sin descripción')})")

    lineas += ["", "DETALLE", ""]
    for fila in incomplete_rows:
        lineas.append(
            f"  [dato_incompleto] fila Excel {fila.get('fila')} | doc: {fila.get('documento') or '(vacío)'} | "
            f"{fila.get('nombre') or '(sin nombre)'} | {fila.get('detalle')}"
        )
    for fallo in fallos:
        lineas.append(
            f"  {fallo['documento']} | {fallo['nombre']} | {fallo['perfil']} | "
            f"motivo: {fallo['motivo']} | {fallo['detalle']}"
        )
        for i in fallo["intentos"]:
            lineas.append(
                f"      intento {i['intento']}: captcha={i['captcha'] or '-'} -> {i['resultado']} ({i['detalle']})"
            )
        if fallo.get("captura"):
            lineas.append(f"      captura: {fallo['captura']}")

    txt_path.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    return csv_path, txt_path


# ---------------------------------------------------------------------------
# Interacción con el usuario
# ---------------------------------------------------------------------------

def choose_profiles():
    print("\nPerfiles para procesar:")
    print("  1. AUXILIAR")
    print("  2. ENFERMERO")
    print("  3. MEDICO")
    print("  4. TODOS")
    option = input("Selecciona opción (1-4) [4]: ").strip() or "4"
    if option == "1":
        return ["AUXILIAR"]
    if option == "2":
        return ["ENFERMERO"]
    if option == "3":
        return ["MEDICO"]
    return None


def choose_captcha_mode() -> str:
    print("\nModo CAPTCHA:")
    print("  1. IA (usa la cadena de proveedores configurada)")
    print("  2. Manual (tú escribes el CAPTCHA en la página)")
    print("  3. Auto (IA y si falla, manual)")
    option = input("Selecciona opción (1-3) [1]: ").strip() or "1"
    if option == "2":
        return "manual_web"
    if option == "3":
        return "auto_web"
    return "ia"


def choose_batch_size(total_pending):
    value = input(f"Cantidad por tanda (0 = todos, máximo {total_pending}) [0]: ").strip() or "0"
    try:
        size = int(value)
    except ValueError:
        size = 0
    return total_pending if size <= 0 else min(size, total_pending)


def build_arg_parser():
    parser = argparse.ArgumentParser(description="Descarga masiva del Carnet Digital de Vacunación (MiVacuna).")
    parser.add_argument("--config", help="Archivo config.json alternativo")
    parser.add_argument("--excel", help="Ruta del Excel de caracterización")
    parser.add_argument("--salida", help="Carpeta base de salida")
    parser.add_argument("--hoja", help="Nombre o índice de la hoja del Excel")
    parser.add_argument("--fila-encabezado", type=int, help="Fila (1-based) donde están los encabezados")
    parser.add_argument("--perfil", choices=["AUXILIAR", "ENFERMERO", "MEDICO", "TODOS"], help="Perfil a procesar")
    parser.add_argument("--tanda", type=int, help="Cantidad de registros en esta corrida")
    parser.add_argument("--modo", choices=["ia", "manual_web", "auto_web"], help="Modo de CAPTCHA")
    parser.add_argument("--headless", action="store_true", help="Ejecutar sin ventana (se usa UA real de Chrome)")
    parser.add_argument("--visible", action="store_true", help="Ejecutar con ventana")
    parser.add_argument("--si", dest="no_prompts", action="store_true", help="No preguntar nada (usa valores por defecto)")
    parser.add_argument("--reiniciar-estado", action="store_true", help="Ignorar el estado y procesar todo de nuevo")
    parser.add_argument("--dry-run", action="store_true", help="Solo leer y validar el Excel, sin abrir el navegador")
    return parser


# ---------------------------------------------------------------------------
# Validación sin navegador
# ---------------------------------------------------------------------------

def dry_run(settings):
    """Lee el Excel con la configuración dada y muestra un resumen."""
    apply_settings(settings)

    selected_profiles = settings.profiles or None
    print("\n" + "=" * 70)
    print("  VALIDACIÓN DE DATOS (sin navegador)")
    print("=" * 70)
    print(f"Excel:        {EXCEL_PATH}")
    print(f"Hoja:         {settings.sheet_name or '(primera)'}")
    print(f"Encabezados:  fila {settings.header_row}")
    print("=" * 70)

    try:
        users, incomplete, stats = load_users(
            EXCEL_PATH,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
            selected_profiles=selected_profiles,
        )
    except Exception as exc:
        print(f"\n[ERROR] {exc}")
        return 2

    print(f"Hoja leída:      {stats['hoja']}")
    print(f"Filas con datos: {stats['total_filas']}")
    print(f"Válidos:         {stats['validos']}")
    print(f"Incompletos:     {stats['incompletos']}")
    print(f"Duplicados:      {stats['duplicados']}")
    print(f"Filtrados:       {stats['filtrados_por_perfil']}")
    print("\nColumnas usadas:")
    for clave, columna in stats["columnas"].items():
        print(f"  - {clave:10s} -> {columna}")

    if incomplete:
        print("\nPrimeros incompletos:")
        for fila in incomplete[:10]:
            print(f"  fila {fila['fila']}: {fila['detalle']}")

    if stats["duplicados_detalle"]:
        print("\nPrimeros duplicados:")
        for fila in stats["duplicados_detalle"][:10]:
            print(f"  fila {fila['fila']}: {fila['documento']} (ya estaba en la fila {fila['primera_fila']})")

    print("\n[OK] Validación terminada.")
    return 0


# ---------------------------------------------------------------------------
# Ejecución
# ---------------------------------------------------------------------------

def run(settings, on_progress=None, on_user=None, should_stop=None):
    """Ejecuta la descarga completa.

    `on_progress(hechos, total, exitosos, fallidos)` y
    `on_user(indice, total, usuario)` permiten que la GUI muestre el avance.
    Devuelve un diccionario con el resultado.
    """
    apply_settings(settings)
    set_stop_checker(should_stop)

    interactive = bool(settings.interactive)
    detenido = False

    print("\n" + "=" * 70)
    print("  DESCARGA AUTOMATICA CARNET DIGITAL VACUNACION - MIVACUNA")
    print("=" * 70)
    print(f"Sistema:      {platform.system()} {platform.machine()}")
    print(f"Excel fuente: {EXCEL_PATH}")
    print(f"Salida base:  {BASE_OUTPUT_DIR}")
    print("=" * 70)

    configure_download_folders()

    if not Path(EXCEL_PATH).exists():
        print(f"\n[ERROR] No existe el Excel: {EXCEL_PATH}")
        print('   Usa --excel "ruta" o elige otro archivo en la interfaz.')
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    # Perfiles
    if settings.profiles is not None:
        selected_profiles = None if not settings.profiles else list(settings.profiles)
    elif not interactive:
        selected_profiles = None
    else:
        selected_profiles = choose_profiles()

    try:
        users, incomplete_rows = load_users_from_excel(
            EXCEL_PATH,
            selected_profiles=selected_profiles,
            sheet_name=settings.sheet_name,
            header_row=settings.header_row,
            column_map=settings.column_map,
        )
    except Exception as exc:
        print(f"\n[ERROR] Error leyendo el Excel: {exc}")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if not users:
        print("\n[ERROR] No se encontraron usuarios validos en el Excel para ese filtro.")
        return {"codigo": 2, "success": 0, "failed": 0, "total": 0, "detenido": False}

    state = load_state()
    processed_set = set(state.get("procesados", []))

    print(f"\nUsuarios validos: {len(users)}")
    print(f"Filas con datos incompletos: {len(incomplete_rows)} (se reportaran como fallos)")
    print(f"Procesados previos en estado: {len(processed_set)}")

    if not settings.reset_state:
        if interactive:
            if settings.resume:
                answer = input("¿Reanudar omitiendo procesados previos? [S/n]: ").strip().lower()
                resume = answer != "n"
            else:
                answer = input("¿Reanudar omitiendo procesados previos? [s/N]: ").strip().lower()
                resume = answer == "s"
        else:
            resume = bool(settings.resume)

        if resume:
            users = [u for u in users if u["documento"] not in processed_set]

    if not users:
        print("\n[OK] No hay pendientes para los perfiles seleccionados.")
        return {"codigo": 0, "success": 0, "failed": 0, "total": 0, "detenido": False}

    if settings.batch_size is not None:
        batch_size = settings.batch_size if settings.batch_size > 0 else len(users)
    elif not interactive:
        batch_size = len(users)
    else:
        batch_size = choose_batch_size(len(users))
    users = users[:batch_size]

    captcha_mode = settings.captcha_mode or ("ia" if not interactive else choose_captcha_mode())

    if settings.headless is not None:
        headless = bool(settings.headless)
    elif not interactive:
        headless = False
    else:
        headless = (input("¿Ejecutar visible? [S/n]: ").strip().lower() or "s") != "s"

    total = len(users)
    print(f"\n[INFO] Esta tanda procesara {total} usuarios. Modo CAPTCHA: {captcha_mode}")

    driver = None
    success = 0
    failed = 0
    fallos = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if on_progress:
        try:
            on_progress(0, total, 0, 0)
        except Exception:
            pass

    try:
        driver = create_driver(headless=headless)
        wait = WebDriverWait(driver, 30)
        warmup_page(driver, wait, refresh_times=1)

        for idx, user in enumerate(users, start=1):
            if stop_requested():
                detenido = True
                print("\n[INFO] Detenido por el usuario. Se generara el informe con lo procesado.")
                break

            user["captcha_mode"] = captcha_mode
            print(
                f"\n[{idx}/{total}] {user['apellido']} {user['nombre']} "
                f"({user['documento']}) - {user['perfil_grupo']}"
            )
            if on_user:
                try:
                    on_user(idx, total, user)
                except Exception:
                    pass

            ok, intentos, motivo, detalle, captura = process_user(
                driver, wait, user, captcha_mode=captcha_mode
            )

            if motivo == "detenido":
                detenido = True
                print("\n[INFO] Detenido por el usuario. Se generara el informe con lo procesado.")
                break

            if ok:
                success += 1
                processed_set.add(user["documento"])
                state.setdefault("fallidos", [])
                if user["documento"] in state["fallidos"]:
                    state["fallidos"].remove(user["documento"])
            else:
                failed += 1
                fallos.append({
                    "documento": user["documento"],
                    "nombre": f"{user['nombre']} {user['apellido']}",
                    "perfil": user["perfil_grupo"],
                    "fila_excel": user.get("fila_excel", ""),
                    "motivo": motivo or "sin_descarga",
                    "detalle": detalle,
                    "intentos": intentos,
                    "captura": captura,
                })
                state.setdefault("fallidos", [])
                if user["documento"] not in state["fallidos"]:
                    state["fallidos"].append(user["documento"])

            state["procesados"] = sorted(processed_set)
            save_state(state)

            if on_progress:
                try:
                    on_progress(idx, total, success, failed)
                except Exception:
                    pass

            time.sleep(0.8)

    except KeyboardInterrupt:
        detenido = True
        print("\n\n[INFO] Proceso interrumpido por el usuario. Se generara el informe con lo procesado.")
    except Exception as exc:
        print(f"\n[ERROR] Error general: {exc}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    csv_path, txt_path = write_failure_report(fallos, incomplete_rows, timestamp)

    print("\n" + "=" * 70)
    print("RESUMEN")
    print(f"[OK] Exitosos: {success}")
    print(f"[ERROR] Fallidos: {failed}")
    print(f"Carpeta base: {BASE_OUTPUT_DIR}")
    print(f"Estado: {STATE_FILE}")
    if txt_path:
        print(f"Informe de fallos: {txt_path}")
        print(f"Informe (CSV):     {csv_path}")
    else:
        print("[OK] No hubo fallos que reportar.")
    print("=" * 70)

    return {
        "codigo": 0,
        "success": success,
        "failed": failed,
        "total": total,
        "detenido": detenido,
        "csv": str(csv_path) if csv_path else "",
        "txt": str(txt_path) if txt_path else "",
        "carpeta": str(BASE_OUTPUT_DIR),
    }


# ---------------------------------------------------------------------------
# main (CLI)
# ---------------------------------------------------------------------------

def main(argv=None):
    args = build_arg_parser().parse_args(argv)

    settings = load_settings(Path(args.config) if args.config else None)

    if args.excel:
        settings.excel_path = Path(args.excel)
    if args.salida:
        settings.output_dir = Path(args.salida)
    if args.hoja is not None:
        settings.sheet_name = args.hoja
    if args.fila_encabezado is not None:
        settings.header_row = max(1, args.fila_encabezado)
    if args.perfil:
        settings.profiles = None if args.perfil == "TODOS" else [args.perfil]
    if args.tanda is not None:
        settings.batch_size = args.tanda
    if args.modo:
        settings.captcha_mode = args.modo
    if args.headless:
        settings.headless = True
    if args.visible:
        settings.headless = False
    if args.reiniciar_estado:
        settings.reset_state = True
    if args.no_prompts:
        settings.interactive = False
        # En la CLI sin preguntas no hay diálogo para el CAPTCHA manual.
        if not args.modo:
            settings.captcha_mode = "ia"

    if args.dry_run:
        return dry_run(settings)

    return run(settings)["codigo"]


if __name__ == "__main__":
    raise SystemExit(main())
