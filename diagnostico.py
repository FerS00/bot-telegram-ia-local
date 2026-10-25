"""Diagnóstico local de ShopIA: libro, token (solo presencia), Ollama y SQLite.

No inicia Telegram y no imprime filas del libro ni el valor de ningún secreto.
Uso: .venv\\Scripts\\python.exe diagnostico.py [--sin-ping]
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

from config import load_settings
from excel_repository import ExcelRepository, WorkbookError
from ollama_client import OllamaClient


def _linea(etiqueta: str, valor: str) -> None:
    print(f"{etiqueta:<22} {valor}")


def _revisar_libro(ruta: Path) -> tuple[bool, ExcelRepository | None]:
    _linea("Libro Excel:", str(ruta.name))
    try:
        repo = ExcelRepository(ruta).load()
    except WorkbookError as exc:
        _linea("  Validación:", f"FALLIDA - {exc}")
        return False, None
    _linea("  Validación:", "OK")
    _linea("  SKU en Productos:", str(len(repo.productos())))
    for hoja, filas in repo.conteos().items():
        _linea(f"  Filas {hoja}:", str(filas))
    return True, repo


async def _revisar_ollama(url: str, modelo: str, timeout_s: float, ping: bool) -> bool:
    cliente = OllamaClient(url, modelo, timeout_s=timeout_s)
    estado = await cliente.estado()
    _linea("Ollama (API):", "disponible" if estado.alcanzable else "no disponible")
    if not estado.alcanzable:
        _linea("  Detalle:", estado.detalle)
        return False
    _linea("  Modelo configurado:", f"{modelo} ({'instalado' if estado.modelo_disponible else 'NO instalado'})")
    if not estado.modelo_disponible:
        _linea("  Detalle:", estado.detalle)
        return False
    if not ping:
        _linea("  Prueba de chat:", "omitida (--sin-ping)")
        return True
    respuesta = await cliente.chat([{"role": "user", "content": "Responde solo: OK"}])
    if respuesta.ok:
        _linea("  Prueba de chat:", "OK")
        return True
    _linea("  Prueba de chat:", f"FALLIDA ({respuesta.error}) - {respuesta.detalle}")
    return False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Diagnóstico de ShopIA (sin Telegram).")
    parser.add_argument("--sin-ping", action="store_true", help="no envía el mensaje de prueba al modelo")
    args = parser.parse_args(argv)

    ajustes = load_settings()
    print("== ShopIA: diagnóstico ==")

    libro_ok, _ = _revisar_libro(ajustes.excel_path)
    _linea("Token Telegram:", "configurado" if ajustes.telegram_token else "NO configurado (solo se revela si existe)")
    _linea("SQLite (runtime):", f"{ajustes.sqlite_path.name} (se crea al arrancar main.py)")

    ollama_ok = asyncio.run(
        _revisar_ollama(ajustes.ollama_url, ajustes.ollama_model, ajustes.ollama_timeout_s, not args.sin_ping)
    )
    if not ollama_ok:
        print("Ollama no listo: el bot funcionará solo con consultas deterministas.")

    if not libro_ok:
        print("Estado global: FALLIDO (el libro es obligatorio).")
        return 1
    if not ajustes.telegram_token:
        print("Estado global: Aprobado con observaciones (falta TELEGRAM_BOT_TOKEN; el bot no arrancará).")
        return 0
    print("Estado global: Aprobado" if ollama_ok else "Estado global: Aprobado con observaciones (Ollama no listo).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
