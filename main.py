"""Punto de entrada del bot de Telegram de ShopIA (long polling, sin webhooks).

Requiere TELEGRAM_BOT_TOKEN. Sin token o con un libro inválido termina con un
mensaje claro. Ollama es opcional: si no responde, el bot sigue operando con
las consultas deterministas.
"""

from __future__ import annotations

import functools
import logging
import sys
from collections.abc import Awaitable, Callable

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters

from config import ConfigError, load_settings, require_telegram_token
from excel_repository import ExcelRepository, WorkbookError
from ollama_client import OllamaClient
from shop_service import AYUDA, ShopError, ShopService
from storage import Storage

logger = logging.getLogger("shopia")

Manejador = Callable[[Update, ContextTypes.DEFAULT_TYPE], Awaitable[None]]


async def _responder(update: Update, texto: str) -> None:
    mensaje = update.effective_message
    if mensaje is not None:
        await mensaje.reply_text(texto)


def _servicio(context: ContextTypes.DEFAULT_TYPE) -> ShopService:
    return context.application.bot_data["shop"]


def _protegido(fn: Manejador) -> Manejador:
    """Convierte errores de negocio en respuestas y oculta cualquier error interno."""

    @functools.wraps(fn)
    async def envoltura(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
        try:
            await fn(update, context)
        except ShopError as exc:
            await _responder(update, str(exc))
        except Exception as exc:  # noqa: BLE001 - nunca dejar caer el bot
            # Solo el tipo de excepción: los mensajes pueden contener URLs con el token.
            logger.error("Error interno en %s (%s)", fn.__name__, type(exc).__name__)
            await _responder(update, "Ocurrió un error interno. Intenta de nuevo.")

    return envoltura


def _texto_args(context: ContextTypes.DEFAULT_TYPE) -> str:
    return " ".join(context.args or [])


def _usuario(update: Update) -> int:
    if update.effective_user is None:
        raise ShopError("No pude identificar tu cuenta de Telegram.")
    return update.effective_user.id


@_protegido
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    teclado = InlineKeyboardMarkup([
        [InlineKeyboardButton("Catálogo", callback_data="cat"), InlineKeyboardButton("Ofertas", callback_data="ofertas")],
        [InlineKeyboardButton("Envíos", callback_data="envios"), InlineKeyboardButton("Carrito", callback_data="carrito")],
    ])
    mensaje = update.effective_message
    if mensaje is not None:
        await mensaje.reply_text(AYUDA, reply_markup=teclado)


@_protegido
async def cmd_catalogo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).catalogo())


@_protegido
async def cmd_buscar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).buscar(_usuario(update), _texto_args(context)))


@_protegido
async def cmd_producto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        raise ShopError("Uso: /producto SKU")
    await _responder(update, _servicio(context).detalle_producto(_usuario(update), context.args[0]))


@_protegido
async def cmd_comparar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args or len(context.args) != 2:
        raise ShopError("Uso: /comparar SKU1 SKU2")
    await _responder(update, _servicio(context).comparar(_usuario(update), *context.args))


@_protegido
async def cmd_agregar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args or len(context.args) > 2:
        raise ShopError("Uso: /agregar SKU [cantidad]")
    cantidad = context.args[1] if len(context.args) == 2 else "1"
    await _responder(update, _servicio(context).agregar(_usuario(update), context.args[0], cantidad))


@_protegido
async def cmd_carrito(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).ver_carrito(_usuario(update)))


@_protegido
async def cmd_quitar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        raise ShopError("Uso: /quitar SKU")
    await _responder(update, _servicio(context).quitar(_usuario(update), context.args[0]))


@_protegido
async def cmd_vaciar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).vaciar(_usuario(update)))


@_protegido
async def cmd_envio(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).envio(_usuario(update), _texto_args(context)))


@_protegido
async def cmd_promos(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).promos())


@_protegido
async def cmd_cupon(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).cupon(_usuario(update), _texto_args(context)))


@_protegido
async def cmd_comprar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await _responder(update, _servicio(context).comprar(_usuario(update)))


@_protegido
async def cmd_pedido(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not context.args:
        raise ShopError("Uso: /pedido CODIGO")
    await _responder(update, _servicio(context).pedido(_usuario(update), context.args[0]))


@_protegido
async def mensaje_libre(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    texto = update.effective_message.text if update.effective_message else ""
    await _responder(update, await _servicio(context).responder(_usuario(update), texto or ""))


_ACCIONES_BOTONES: dict[str, str] = {
    "cat": "/catalogo",
    "ofertas": "/promos",
    "carrito": "/carrito",
}


@_protegido
async def boton(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    consulta = update.callback_query
    if consulta is None:
        return
    await consulta.answer()
    accion = _ACCIONES_BOTONES.get(consulta.data or "")
    servicio = _servicio(context)
    usuario = consulta.from_user.id
    if accion == "/catalogo":
        texto = servicio.catalogo()
    elif accion == "/promos":
        texto = servicio.promos()
    elif accion == "/carrito":
        texto = servicio.ver_carrito(usuario)
    else:
        texto = "Para cotizar envío escribe /envio DISTRITO, por ejemplo: /envio Los Olivos."
    await _responder(update, texto)


async def _al_iniciar(application: Application) -> None:
    estado = await application.bot_data["ollama"].estado()
    if estado.alcanzable and estado.modelo_disponible:
        logger.info("Ollama disponible; modelo configurado instalado.")
    else:
        logger.warning("Ollama no listo (%s). El bot funciona en modo determinista.", estado.detalle or "sin detalle")


async def _error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error("Error no controlado en el bot (%s)", type(context.error).__name__)


def construir_aplicacion(token: str, servicio: ShopService, ollama: OllamaClient) -> Application:
    app = Application.builder().token(token).post_init(_al_iniciar).build()
    app.bot_data["shop"] = servicio
    app.bot_data["ollama"] = ollama
    comandos = {
        "start": cmd_start,
        "catalogo": cmd_catalogo,
        "buscar": cmd_buscar,
        "producto": cmd_producto,
        "comparar": cmd_comparar,
        "agregar": cmd_agregar,
        "carrito": cmd_carrito,
        "quitar": cmd_quitar,
        "vaciar": cmd_vaciar,
        "envio": cmd_envio,
        "promos": cmd_promos,
        "cupon": cmd_cupon,
        "comprar": cmd_comprar,
        "pedido": cmd_pedido,
    }
    for nombre, manejador in comandos.items():
        app.add_handler(CommandHandler(nombre, manejador))
    app.add_handler(CallbackQueryHandler(boton))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, mensaje_libre))
    app.add_error_handler(_error)
    return app


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    # httpx registra cada URL, que contiene el token de Telegram: lo silenciamos.
    for nombre in ("httpx", "httpcore"):
        logging.getLogger(nombre).setLevel(logging.WARNING)

    try:
        ajustes = load_settings()
        token = require_telegram_token(ajustes)
    except ConfigError as exc:
        print(f"Error de configuración: {exc}", file=sys.stderr)
        return 2

    try:
        repo = ExcelRepository(ajustes.excel_path).load()
    except WorkbookError as exc:
        print(f"Error en el libro Excel: {exc}", file=sys.stderr)
        return 3

    try:
        ollama = OllamaClient(ajustes.ollama_url, ajustes.ollama_model, timeout_s=ajustes.ollama_timeout_s)
    except ConfigError as exc:
        print(f"Error de configuración: {exc}", file=sys.stderr)
        return 2
    servicio = ShopService(repo, Storage(ajustes.sqlite_path), ollama)

    logger.info("ShopIA iniciada con %d SKU. Long polling activo.", len(repo.productos()))
    aplicacion = construir_aplicacion(token, servicio, ollama)
    aplicacion.run_polling(drop_pending_updates=False)  # conserva los mensajes pendientes del usuario
    return 0


if __name__ == "__main__":
    sys.exit(main())
