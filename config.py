"""Configuración de ShopIA.

Lee variables de entorno y el archivo .env de la raíz del proyecto. Importar
este módulo no exige token ni Ollama; los errores se reportan al arrancar.
"""

from __future__ import annotations

import os
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import urlparse

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = BASE_DIR / "ShopIA_Base_Datos_RAG.xlsx"
DEFAULT_SQLITE_PATH = BASE_DIR / "runtime" / "shopia.sqlite3"
DEFAULT_OLLAMA_URL = "http://127.0.0.1:11434"
DEFAULT_OLLAMA_MODEL = "qwen2.5-coder:7b"
DEFAULT_OLLAMA_TIMEOUT_S = 120.0
LOOPBACK_HOSTS = frozenset({"127.0.0.1", "localhost", "::1"})
TOKEN_PLACEHOLDER_PREFIX = "<"


class ConfigError(RuntimeError):
    """Error de configuración con un mensaje apto para mostrar al usuario."""


@dataclass(frozen=True)
class Settings:
    excel_path: Path
    sqlite_path: Path
    ollama_url: str
    ollama_model: str
    ollama_timeout_s: float
    telegram_token: str | None = field(default=None, repr=False)


def _ruta(valor: str | None, por_defecto: Path) -> Path:
    if not valor or not valor.strip():
        return por_defecto
    ruta = Path(valor.strip())
    return ruta if ruta.is_absolute() else BASE_DIR / ruta


def _url_local(valor: str) -> str:
    url = valor.strip().rstrip("/")
    partes = urlparse(url)
    if partes.scheme not in {"http", "https"} or partes.hostname not in LOOPBACK_HOSTS:
        raise ConfigError(
            "OLLAMA_BASE_URL debe apuntar a localhost (127.0.0.1). "
            "ShopIA no usa proveedores de IA externos."
        )
    return url


def _token(valor: str | None) -> str | None:
    if valor is None:
        return None
    limpio = valor.strip()
    # Un marcador de .env.example no es un token real.
    if not limpio or limpio.startswith(TOKEN_PLACEHOLDER_PREFIX):
        return None
    return limpio


def leer_settings(entorno: Mapping[str, str]) -> Settings:
    """Construye Settings a partir de un mapa de variables (sin tocar disco)."""
    try:
        timeout = float(entorno.get("OLLAMA_TIMEOUT_S") or DEFAULT_OLLAMA_TIMEOUT_S)
    except ValueError as exc:
        raise ConfigError("OLLAMA_TIMEOUT_S debe ser un número de segundos.") from exc
    if timeout <= 0:
        raise ConfigError("OLLAMA_TIMEOUT_S debe ser mayor que cero.")

    return Settings(
        excel_path=_ruta(entorno.get("EXCEL_PATH"), DEFAULT_EXCEL_PATH),
        sqlite_path=_ruta(entorno.get("SQLITE_PATH"), DEFAULT_SQLITE_PATH),
        ollama_url=_url_local(entorno.get("OLLAMA_BASE_URL") or DEFAULT_OLLAMA_URL),
        ollama_model=(entorno.get("OLLAMA_MODEL") or DEFAULT_OLLAMA_MODEL).strip(),
        ollama_timeout_s=timeout,
        telegram_token=_token(entorno.get("TELEGRAM_BOT_TOKEN")),
    )


def load_settings() -> Settings:
    """Carga .env de la raíz (si existe) y las variables de entorno."""
    load_dotenv(BASE_DIR / ".env", override=False)
    return leer_settings(os.environ)


def require_telegram_token(settings: Settings) -> str:
    """Devuelve el token o lanza ConfigError sin revelar ningún valor."""
    if not settings.telegram_token:
        raise ConfigError(
            "Falta TELEGRAM_BOT_TOKEN. Completa TELEGRAM_BOT_TOKEN en el archivo .env "
            "de la raíz del proyecto con tu token de BotFather (conserva las demás "
            "variables) y vuelve a ejecutar main.py."
        )
    return settings.telegram_token
