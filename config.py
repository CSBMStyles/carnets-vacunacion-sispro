"""Configuración persistente del proyecto.

Dos capas, a propósito:

- ``config.json``: rutas, hoja, mapeo de columnas y comportamiento. No tiene
  secretos, así que se puede copiar o compartir.
- ``.env``: solo API keys y ajustes de IA. Nunca se comparte.

Los valores de ``config.json`` mandan sobre los del ``.env`` para las rutas,
porque la interfaz guarda ahí lo que el usuario elige.
"""

from __future__ import annotations

import json
import os
import platform
import re
import shutil
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

try:  # python-dotenv es opcional: la app funciona sin él.
    from dotenv import load_dotenv
except Exception:  # pragma: no cover
    load_dotenv = None


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = PROJECT_DIR / "config.json"
ENV_FILE = PROJECT_DIR / ".env"

# Campos que solo viven durante una ejecución y no se guardan.
_RUNTIME_FIELDS = {"interactive", "reset_state"}

# Rutas portables: las internas se guardan relativas al proyecto y las de
# datos (Excel/salida) relativas a la carpeta Vacunaciones.
WORKSPACE_DIR = PROJECT_DIR.parents[2] if len(PROJECT_DIR.parents) > 2 else PROJECT_DIR
_PROJECT_FIELDS = {"state_file", "temp_dir", "error_dir"}
_PORTABLE_FIELDS = {"excel_path", "output_dir"} | _PROJECT_FIELDS


def _safe_parent(index: int) -> Path | None:
    parents = PROJECT_DIR.parents
    return parents[index] if len(parents) > index else None


def default_excel_path() -> Path:
    """Busca la caracterización en las ubicaciones conocidas.

    Preferencia:
      1. ``Codigo/CARACTERIZACION.xlsx`` (base vigente).
      2. ``CARACTERIZACION.xlsx`` en la carpeta del proyecto.
      3. ``CARACTERIZACION COPIA.xlsx`` (lote anterior).
    """
    candidates: list[Path] = []

    codigo = _safe_parent(1)
    if codigo:
        candidates.append(codigo / "CARACTERIZACION.xlsx")

    vacunaciones = _safe_parent(2)
    if vacunaciones:
        candidates.append(vacunaciones / "CARACTERIZACION.xlsx")
        candidates.append(vacunaciones / "CARACTERIZACION COPIA.xlsx")

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return candidates[0] if candidates else PROJECT_DIR / "CARACTERIZACION.xlsx"


def default_output_dir() -> Path:
    vacunaciones = _safe_parent(2)
    if vacunaciones:
        candidate = vacunaciones / "CARNETS DIGITAL ACTUALIZADO"
        if candidate.exists():
            return candidate
    return PROJECT_DIR / "carnets"


def _as_path(value: Any, fallback: Path) -> Path:
    if value in (None, ""):
        return fallback
    return Path(os.path.expandvars(os.path.expanduser(str(value))))


@dataclass
class Settings:
    # Entrada / salida
    excel_path: Path = field(default_factory=default_excel_path)
    output_dir: Path = field(default_factory=default_output_dir)
    state_file: Path = field(default_factory=lambda: PROJECT_DIR / "estado_proceso_actualizado.json")
    temp_dir: Path = field(default_factory=lambda: PROJECT_DIR / "temp_pdf")
    error_dir: Path = field(default_factory=lambda: PROJECT_DIR / "errors")

    # Lectura del Excel
    sheet_name: str | int | None = None
    header_row: int = 1
    column_map: dict[str, str] = field(default_factory=dict)

    # Comportamiento
    profiles: list[str] | None = None
    batch_size: int | None = None
    captcha_mode: str | None = None
    headless: bool | None = None
    resume: bool = True
    reset_state: bool = False
    captcha_max_attempts: int = 3
    max_image_retries: int = 4
    download_timeout: int = 25
    driver_backend: str = "auto"
    captcha_providers: str = ""

    # Solo runtime
    interactive: bool = True

    # ------------------------------------------------------------------
    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        for key in _RUNTIME_FIELDS:
            data.pop(key, None)
        for key, value in list(data.items()):
            if isinstance(value, Path):
                if key in _PORTABLE_FIELDS:
                    base = PROJECT_DIR if key in _PROJECT_FIELDS else WORKSPACE_DIR
                    try:
                        data[key] = str(value.relative_to(base))
                        continue
                    except ValueError:
                        pass
                data[key] = str(value)
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "Settings":
        settings = cls()
        valid = {f.name for f in fields(cls)}
        for key, value in (data or {}).items():
            if key not in valid:
                continue

            if key in _PORTABLE_FIELDS:
                fallback = getattr(settings, key)
                raw = str(value)

                if raw and not Path(raw).is_absolute():
                    base = PROJECT_DIR if key in _PROJECT_FIELDS else WORKSPACE_DIR
                    candidate = base / raw
                else:
                    candidate = _as_path(value, fallback)

                # Si el Excel guardado ya no existe (otra máquina), se usa el
                # que exista por defecto.
                if key == "excel_path" and not candidate.exists():
                    candidate = fallback

                # Carpetas internas de otra máquina: se recrean en el proyecto.
                if key in _PROJECT_FIELDS and not candidate.exists():
                    try:
                        inside = candidate.resolve().is_relative_to(PROJECT_DIR.resolve())
                    except Exception:
                        inside = str(candidate).startswith(str(PROJECT_DIR))
                    if not inside:
                        candidate = fallback

                setattr(settings, key, candidate)
            elif key == "header_row":
                try:
                    settings.header_row = max(1, int(value))
                except (TypeError, ValueError):
                    pass
            elif key in ("captcha_max_attempts", "max_image_retries", "download_timeout"):
                try:
                    setattr(settings, key, max(1, int(value)))
                except (TypeError, ValueError):
                    pass
            elif key == "batch_size":
                try:
                    settings.batch_size = None if value is None else max(0, int(value))
                except (TypeError, ValueError):
                    pass
            elif key == "profiles":
                if value in (None, "", []):
                    settings.profiles = None
                elif isinstance(value, str):
                    settings.profiles = [value]
                else:
                    settings.profiles = [str(v) for v in value]
            elif key == "column_map" and isinstance(value, dict):
                settings.column_map = {str(k): str(v) for k, v in value.items() if v}
            else:
                setattr(settings, key, value)

        return settings


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return max(1, int(raw))
    except ValueError:
        return default


def load_settings(path: Path | None = None) -> Settings:
    path = Path(path) if path else CONFIG_FILE
    if not path.exists():
        settings = Settings()
        # Compatibilidad: en la primera ejecución se respetan los valores del .env.
        # Después manda config.json (lo que se elige en la interfaz).
        settings.captcha_max_attempts = _env_int("CAPTCHA_MAX_ATTEMPTS", settings.captcha_max_attempts)
        settings.max_image_retries = _env_int("MAX_IMAGE_RETRIES", settings.max_image_retries)
        settings.download_timeout = _env_int("DOWNLOAD_TIMEOUT", settings.download_timeout)
        backend = os.getenv("DRIVER_BACKEND", "").strip()
        if backend:
            settings.driver_backend = backend
        providers = os.getenv("CAPTCHA_PROVIDERS", "").strip()
        if providers:
            settings.captcha_providers = providers
        return settings

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return Settings()

    return Settings.from_dict(data)


def save_settings(settings: Settings, path: Path | None = None) -> Path:
    path = Path(path) if path else CONFIG_FILE
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(settings.to_dict(), ensure_ascii=False, indent=2)

    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(payload + "\n", encoding="utf-8")
    os.replace(tmp, path)
    return path


# ----------------------------------------------------------------------
# .env
# ----------------------------------------------------------------------

def load_environment() -> None:
    """Carga el .env del proyecto sin sobrescribir variables ya definidas."""
    if load_dotenv is None:
        return
    if ENV_FILE.exists():
        load_dotenv(ENV_FILE, override=False)


_ENV_LINE = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")


def read_env_file(path: Path | None = None) -> dict[str, str]:
    path = Path(path) if path else ENV_FILE
    values: dict[str, str] = {}
    if not path.exists():
        return values

    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        match = _ENV_LINE.match(line)
        if not match:
            continue
        key, raw = match.group(1), match.group(2).strip()
        if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ("'", '"'):
            raw = raw[1:-1]
        values[key] = raw
    return values


def _env_value(value: str) -> str:
    value = str(value or "")
    if not value:
        return ""
    if any(ch in value for ch in (" ", "#", "\t")) or value != value.strip():
        escaped = value.replace('"', '\\"')
        return f'"{escaped}"'
    return value


def update_env_file(values: dict[str, str], path: Path | None = None) -> Path:
    """Actualiza claves del .env conservando comentarios y el resto de líneas."""
    path = Path(path) if path else ENV_FILE
    path.parent.mkdir(parents=True, exist_ok=True)

    lines = path.read_text(encoding="utf-8").splitlines() if path.exists() else []
    pending = {k: str(v) for k, v in values.items()}
    output: list[str] = []

    for line in lines:
        match = _ENV_LINE.match(line)
        if match and match.group(1) in pending:
            key = match.group(1)
            output.append(f"{key}={_env_value(pending.pop(key))}")
        else:
            output.append(line)

    for key, value in pending.items():
        output.append(f"{key}={_env_value(value)}")

    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text("\n".join(output).rstrip() + "\n", encoding="utf-8")
    os.replace(tmp, path)
    return path


def mask_secret(value: str, visible: int = 4) -> str:
    value = str(value or "")
    if not value:
        return ""
    if len(value) <= visible * 2:
        return "•" * len(value)
    return f"{value[:visible]}…{value[-visible:]}"


# ----------------------------------------------------------------------
# Entorno
# ----------------------------------------------------------------------

def chrome_path() -> Path | None:
    """Devuelve la ruta de Chrome si está instalado, o None."""
    system = platform.system()

    candidates: list[Path] = []
    if system == "Windows":
        for env in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
            base = os.environ.get(env)
            if base:
                candidates.append(Path(base) / "Google" / "Chrome" / "Application" / "chrome.exe")
    elif system == "Darwin":
        candidates.append(Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"))
        candidates.append(Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
    else:
        for name in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            found = shutil.which(name)
            if found:
                candidates.append(Path(found))

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


load_environment()
