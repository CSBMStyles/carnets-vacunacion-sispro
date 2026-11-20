"""Utilidades de texto, fechas y nombres de archivo.

Se separaron de main.py para poder reutilizarlas en la interfaz y en los
demás proyectos de antecedentes sin arrastrar Selenium.
"""

from __future__ import annotations

import sys
import unicodedata
from datetime import datetime

import pandas as pd


def configure_console() -> None:
    """Evita UnicodeEncodeError en la consola de Windows (cp1252)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# Caracteres que Windows no permite en un nombre de archivo.
INVALID_FILENAME_CHARS = '<>:"/\\|?*'

# Nombres reservados por Windows (sin extensión). Si el nombre coincide,
# Windows lo rechaza aunque tenga extensión.
RESERVED_WINDOWS_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}


def normalize_text(value) -> str:
    """Mayúsculas, sin acentos y con espacios colapsados."""
    text = str(value or "").strip().upper()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    return " ".join(text.split())


def safe_filename(text: str, max_stem: int = 120) -> str:
    """Devuelve un nombre seguro para macOS y Windows.

    - Quita acentos (evita el problema NFD/NFC de macOS).
    - Reemplaza caracteres inválidos de Windows.
    - Recorta a `max_stem` caracteres para no superar MAX_PATH.
    - Quita puntos/espacios finales y evita nombres reservados (CON, PRN...).
    """
    text = str(text or "").strip()
    text = "".join(c for c in unicodedata.normalize("NFD", text) if unicodedata.category(c) != "Mn")
    cleaned = "".join(
        "_" if (c in INVALID_FILENAME_CHARS or ord(c) < 32) else c for c in text
    )
    cleaned = " ".join(cleaned.split()).strip(" .")

    if max_stem > 0 and len(cleaned) > max_stem:
        cleaned = cleaned[:max_stem].rstrip(" .")

    if cleaned.upper() in RESERVED_WINDOWS_NAMES:
        cleaned = f"_{cleaned}"

    return cleaned or "sin_nombre"


def classify_profile(raw_profile) -> str | None:
    """Mapea el texto libre de 'Perfil' o 'Cargo' a uno de los 3 grupos."""
    profile = normalize_text(raw_profile)
    if "AUXILIAR" in profile:
        return "AUXILIAR"
    if "ENFERM" in profile:
        return "ENFERMERO"
    if "MEDIC" in profile:
        return "MEDICO"
    return None


def clean_document(doc_value) -> str:
    """Devuelve el número de documento solo con dígitos.

    El Excel puede entregarlo como número (1000000001), como float
    (1000000001.0) o como texto ("1000000001.0"); en todos los casos se
    normaliza sin perder ni agregar dígitos.
    """
    if doc_value is None:
        return ""
    if not isinstance(doc_value, str) and pd.isna(doc_value):
        return ""

    if isinstance(doc_value, (int, float)):
        try:
            if float(doc_value).is_integer():
                return str(int(doc_value))
        except Exception:
            pass
    elif hasattr(doc_value, "item"):  # numpy.int64 / numpy.float64
        try:
            return str(int(doc_value))
        except Exception:
            pass

    text = str(doc_value).strip().replace(" ", "")
    if text.endswith(".0"):
        text = text[:-2]
    return "".join(ch for ch in text if ch.isdigit())


def parse_date_for_form(value) -> str:
    """Convierte la fecha del Excel al formato AAAA-MM-DD del formulario."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if hasattr(value, "to_pydatetime"):
        return value.to_pydatetime().strftime("%Y-%m-%d")

    text = str(value).strip()
    if not text:
        return ""

    if len(text) >= 10 and text[0:4].isdigit() and text[4] in ("/", "-"):
        return text.replace("/", "-")[:10]

    parsed = pd.to_datetime(text, dayfirst=False, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m-%d")

    return text.replace("/", "-")
