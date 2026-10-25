"""Cliente de Ollama local (httpx). Nunca lanza excepciones de red al llamador.

Solo habla con /api/tags y /api/chat en localhost. No usa function calling ni
ningún otro proveedor de IA.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
from urllib.parse import urlparse

import httpx

from config import LOOPBACK_HOSTS, ConfigError

CONNECT_TIMEOUT_S = 5.0
TAGS_TIMEOUT_S = 5.0
OPCIONES_CHAT = {"temperature": 0.2, "num_predict": 350, "num_ctx": 4096}


@dataclass(frozen=True)
class EstadoOllama:
    alcanzable: bool
    modelos: tuple[str, ...] = ()
    modelo_disponible: bool = False
    detalle: str = ""


@dataclass(frozen=True)
class RespuestaChat:
    ok: bool
    texto: str = ""
    error: str | None = None  # "no_disponible", "timeout", "modelo_no_encontrado", "http", "respuesta_invalida"
    detalle: str = field(default="", compare=False)


def validar_url_local(url: str) -> str:
    partes = urlparse(url)
    if partes.scheme not in {"http", "https"} or partes.hostname not in LOOPBACK_HOSTS:
        raise ConfigError("El cliente de Ollama solo acepta URLs locales (127.0.0.1 o localhost).")
    return url.rstrip("/")


class OllamaClient:
    def __init__(
        self,
        base_url: str,
        model: str,
        timeout_s: float = 120.0,
        transport: httpx.AsyncBaseTransport | None = None,
    ) -> None:
        self._base = validar_url_local(base_url)
        self._model = model
        self._timeout = httpx.Timeout(timeout_s, connect=CONNECT_TIMEOUT_S)
        self._transport = transport

    @property
    def modelo(self) -> str:
        return self._model

    def _cliente(self, timeout: httpx.Timeout) -> httpx.AsyncClient:
        return httpx.AsyncClient(base_url=self._base, timeout=timeout, transport=self._transport)

    async def estado(self) -> EstadoOllama:
        """GET /api/tags: indica si Ollama responde y si el modelo está instalado."""
        try:
            async with self._cliente(httpx.Timeout(TAGS_TIMEOUT_S)) as cliente:
                resp = await cliente.get("/api/tags")
                resp.raise_for_status()
                datos = resp.json()
        except httpx.TimeoutException:
            return EstadoOllama(False, detalle="Ollama no respondió a tiempo.")
        except httpx.HTTPError:
            return EstadoOllama(False, detalle="Ollama no está disponible en la URL configurada.")
        except ValueError:
            return EstadoOllama(False, detalle="Respuesta no válida de Ollama.")
        if not isinstance(datos, dict):
            return EstadoOllama(False, detalle="Respuesta no válida de Ollama.")

        nombres = tuple(
            str(m.get("name") or m.get("model") or "")
            for m in (datos.get("models") or [])
            if isinstance(m, dict)
        )
        presente = self._model in nombres
        detalle = "" if presente else f"El modelo {self._model} no aparece instalado en Ollama."
        return EstadoOllama(True, nombres, presente, detalle)

    async def chat(self, mensajes: list[dict[str, str]]) -> RespuestaChat:
        """POST /api/chat sin herramientas. Devuelve un resultado; no lanza errores de red."""
        carga = {
            "model": self._model,
            "messages": mensajes,
            "stream": False,
            "options": dict(OPCIONES_CHAT),
        }
        try:
            async with self._cliente(self._timeout) as cliente:
                resp = await cliente.post("/api/chat", json=carga)
        except httpx.ConnectError:
            return RespuestaChat(False, error="no_disponible", detalle="No se pudo conectar con Ollama.")
        except httpx.TimeoutException:
            return RespuestaChat(
                False,
                error="timeout",
                detalle="Ollama no respondió a tiempo. Si es la primera consulta, el modelo puede estar cargándose; reintenta.",
            )
        except httpx.HTTPError:
            return RespuestaChat(False, error="no_disponible", detalle="Error de red con Ollama.")

        if resp.status_code == 404:
            return RespuestaChat(False, error="modelo_no_encontrado", detalle="Modelo no instalado en Ollama.")
        if resp.status_code >= 400:
            return RespuestaChat(False, error="http", detalle=f"Ollama respondió HTTP {resp.status_code}.")

        try:
            datos: Any = resp.json()
            texto = datos["message"]["content"]
        except (ValueError, KeyError, TypeError):
            return RespuestaChat(False, error="respuesta_invalida", detalle="Respuesta no válida de Ollama.")
        if not isinstance(texto, str) or not texto.strip():
            return RespuestaChat(False, error="respuesta_invalida", detalle="Ollama devolvió una respuesta vacía.")
        return RespuestaChat(True, texto=texto.strip())
