"""Persistencia local en SQLite: carritos, preferencias, pedidos SIM- e historial.

Todas las consultas van parametrizadas y las escrituras son transacciones.
Nunca escribe en el libro Excel. El consumo de stock es simulado: se deriva de
las líneas de pedidos SIM- guardados aquí.
"""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator, Mapping, Sequence
from contextlib import contextmanager, closing
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

MAX_HISTORIAL = 8

_ESQUEMA = """
CREATE TABLE IF NOT EXISTS carrito (
    user_id     INTEGER NOT NULL,
    sku         TEXT    NOT NULL,
    cantidad    INTEGER NOT NULL CHECK (cantidad > 0),
    actualizado TEXT    NOT NULL,
    PRIMARY KEY (user_id, sku)
);
CREATE TABLE IF NOT EXISTS preferencias_usuario (
    user_id  INTEGER PRIMARY KEY,
    zona_id  TEXT,
    cupon    TEXT
);
CREATE TABLE IF NOT EXISTS pedido (
    pedido_id     TEXT PRIMARY KEY,
    user_id       INTEGER NOT NULL,
    creado        TEXT    NOT NULL,
    zona_id       TEXT    NOT NULL,
    cupon         TEXT,
    subtotal_pen  TEXT    NOT NULL,
    descuento_pen TEXT    NOT NULL,
    envio_pen     TEXT    NOT NULL,
    total_pen     TEXT    NOT NULL,
    estado        TEXT    NOT NULL
);
CREATE TABLE IF NOT EXISTS pedido_linea (
    pedido_id           TEXT    NOT NULL REFERENCES pedido(pedido_id) ON DELETE CASCADE,
    sku                 TEXT    NOT NULL,
    cantidad            INTEGER NOT NULL CHECK (cantidad > 0),
    precio_unitario_pen TEXT    NOT NULL,
    subtotal_pen        TEXT    NOT NULL,
    PRIMARY KEY (pedido_id, sku)
);
CREATE INDEX IF NOT EXISTS ix_pedido_linea_sku ON pedido_linea (sku);
CREATE TABLE IF NOT EXISTS historial (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id   INTEGER NOT NULL,
    intencion TEXT    NOT NULL,
    skus      TEXT    NOT NULL,
    creado    TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS ix_historial_usuario ON historial (user_id, id);
"""


class StorageError(RuntimeError):
    """Error de persistencia con mensaje apto para el usuario."""


class StockInsuficiente(StorageError):
    def __init__(self, sku: str, disponible: int) -> None:
        super().__init__(f"Stock insuficiente para {sku}: disponibles {disponible}.")
        self.sku = sku
        self.disponible = disponible


@dataclass(frozen=True)
class LineaPedido:
    sku: str
    cantidad: int
    precio_unitario: str
    subtotal: str


@dataclass(frozen=True)
class PedidoRegistrado:
    pedido_id: str
    creado: str
    zona_id: str
    cupon: str | None
    subtotal: str
    descuento: str
    envio: str
    total: str
    estado: str
    lineas: tuple[LineaPedido, ...]


def ahora_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


class Storage:
    def __init__(self, ruta: Path | str) -> None:
        self._ruta = Path(ruta)
        self._ruta.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._conectar()) as conn:
            conn.executescript(_ESQUEMA)

    @property
    def ruta(self) -> Path:
        return self._ruta

    def _conectar(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._ruta, timeout=10, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    @contextmanager
    def _transaccion(self) -> Iterator[sqlite3.Connection]:
        conn = self._conectar()
        try:
            conn.execute("BEGIN IMMEDIATE")
            yield conn
            conn.execute("COMMIT")
        except BaseException:
            if conn.in_transaction:
                conn.execute("ROLLBACK")
            raise
        finally:
            conn.close()

    @staticmethod
    def _consumido(conn: sqlite3.Connection, sku: str) -> int:
        fila = conn.execute(
            "SELECT COALESCE(SUM(cantidad), 0) AS total FROM pedido_linea WHERE sku = ?",
            (sku,),
        ).fetchone()
        return int(fila["total"])

    # ------------------------------------------------------------ carrito

    def agregar_al_carrito(self, user_id: int, sku: str, cantidad: int, stock_total: int) -> int:
        """Suma unidades al carrito si caben en el stock menos lo ya simulado vendido."""
        with self._transaccion() as conn:
            disponible = stock_total - self._consumido(conn, sku)
            fila = conn.execute(
                "SELECT cantidad FROM carrito WHERE user_id = ? AND sku = ?", (user_id, sku)
            ).fetchone()
            actual = int(fila["cantidad"]) if fila else 0
            nueva = actual + cantidad
            if nueva > disponible:
                raise StockInsuficiente(sku, max(disponible, 0))
            conn.execute(
                """
                INSERT INTO carrito (user_id, sku, cantidad, actualizado) VALUES (?, ?, ?, ?)
                ON CONFLICT (user_id, sku) DO UPDATE SET
                    cantidad = excluded.cantidad, actualizado = excluded.actualizado
                """,
                (user_id, sku, nueva, ahora_iso()),
            )
            return nueva

    def lineas_carrito(self, user_id: int) -> list[tuple[str, int]]:
        with closing(self._conectar()) as conn:
            filas = conn.execute(
                "SELECT sku, cantidad FROM carrito WHERE user_id = ? ORDER BY sku", (user_id,)
            ).fetchall()
        return [(f["sku"], int(f["cantidad"])) for f in filas]

    def quitar_del_carrito(self, user_id: int, sku: str) -> bool:
        with self._transaccion() as conn:
            cursor = conn.execute("DELETE FROM carrito WHERE user_id = ? AND sku = ?", (user_id, sku))
            return cursor.rowcount > 0

    def vaciar_carrito(self, user_id: int) -> int:
        with self._transaccion() as conn:
            cursor = conn.execute("DELETE FROM carrito WHERE user_id = ?", (user_id,))
            return cursor.rowcount

    # ------------------------------------------------------------ preferencias

    def preferencias(self, user_id: int) -> tuple[str | None, str | None]:
        with closing(self._conectar()) as conn:
            fila = conn.execute(
                "SELECT zona_id, cupon FROM preferencias_usuario WHERE user_id = ?", (user_id,)
            ).fetchone()
        return (fila["zona_id"], fila["cupon"]) if fila else (None, None)

    def fijar_zona(self, user_id: int, zona_id: str) -> None:
        with self._transaccion() as conn:
            conn.execute(
                """
                INSERT INTO preferencias_usuario (user_id, zona_id) VALUES (?, ?)
                ON CONFLICT (user_id) DO UPDATE SET zona_id = excluded.zona_id
                """,
                (user_id, zona_id),
            )

    def fijar_cupon(self, user_id: int, cupon: str | None) -> None:
        with self._transaccion() as conn:
            conn.execute(
                """
                INSERT INTO preferencias_usuario (user_id, cupon) VALUES (?, ?)
                ON CONFLICT (user_id) DO UPDATE SET cupon = excluded.cupon
                """,
                (user_id, cupon),
            )

    # ------------------------------------------------------------ pedidos

    def confirmar_pedido(
        self,
        user_id: int,
        pedido_id: str,
        zona_id: str,
        cupon: str | None,
        lineas: Sequence[LineaPedido],
        totales: Mapping[str, str],
        stock_total: Mapping[str, int],
    ) -> None:
        """Guarda el pedido SIM-, consume stock simulado y vacía el carrito en una transacción."""
        with self._transaccion() as conn:
            for linea in lineas:
                disponible = stock_total.get(linea.sku, 0) - self._consumido(conn, linea.sku)
                if linea.cantidad > disponible:
                    raise StockInsuficiente(linea.sku, max(disponible, 0))
            conn.execute(
                """
                INSERT INTO pedido (pedido_id, user_id, creado, zona_id, cupon,
                                    subtotal_pen, descuento_pen, envio_pen, total_pen, estado)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Simulado (sin cobro)')
                """,
                (
                    pedido_id, user_id, ahora_iso(), zona_id, cupon,
                    totales["subtotal"], totales["descuento"], totales["envio"], totales["total"],
                ),
            )
            conn.executemany(
                """
                INSERT INTO pedido_linea (pedido_id, sku, cantidad, precio_unitario_pen, subtotal_pen)
                VALUES (?, ?, ?, ?, ?)
                """,
                [(pedido_id, l.sku, l.cantidad, l.precio_unitario, l.subtotal) for l in lineas],
            )
            conn.execute("DELETE FROM carrito WHERE user_id = ?", (user_id,))

    def pedido_de_usuario(self, user_id: int, pedido_id: str) -> PedidoRegistrado | None:
        with closing(self._conectar()) as conn:
            cab = conn.execute(
                "SELECT * FROM pedido WHERE pedido_id = ? AND user_id = ?", (pedido_id, user_id)
            ).fetchone()
            if cab is None:
                return None
            filas = conn.execute(
                "SELECT sku, cantidad, precio_unitario_pen, subtotal_pen FROM pedido_linea "
                "WHERE pedido_id = ? ORDER BY sku",
                (pedido_id,),
            ).fetchall()
        return PedidoRegistrado(
            pedido_id=cab["pedido_id"],
            creado=cab["creado"],
            zona_id=cab["zona_id"],
            cupon=cab["cupon"],
            subtotal=cab["subtotal_pen"],
            descuento=cab["descuento_pen"],
            envio=cab["envio_pen"],
            total=cab["total_pen"],
            estado=cab["estado"],
            lineas=tuple(
                LineaPedido(f["sku"], int(f["cantidad"]), f["precio_unitario_pen"], f["subtotal_pen"])
                for f in filas
            ),
        )

    # ------------------------------------------------------------ historial

    def registrar_historial(self, user_id: int, intencion: str, skus: Sequence[str]) -> None:
        """Guarda solo la intención y los SKU mencionados; nunca el texto libre del usuario."""
        with self._transaccion() as conn:
            conn.execute(
                "INSERT INTO historial (user_id, intencion, skus, creado) VALUES (?, ?, ?, ?)",
                (user_id, intencion[:30], ",".join(skus)[:500], ahora_iso()),
            )
            conn.execute(
                """
                DELETE FROM historial WHERE user_id = ? AND id NOT IN (
                    SELECT id FROM historial WHERE user_id = ? ORDER BY id DESC LIMIT ?
                )
                """,
                (user_id, user_id, MAX_HISTORIAL),
            )

    def historial(self, user_id: int) -> list[tuple[str, tuple[str, ...]]]:
        """Hasta MAX_HISTORIAL intercambios del usuario, del más reciente al más antiguo."""
        with closing(self._conectar()) as conn:
            filas = conn.execute(
                "SELECT intencion, skus FROM historial WHERE user_id = ? ORDER BY id DESC LIMIT ?",
                (user_id, MAX_HISTORIAL),
            ).fetchall()
        return [
            (f["intencion"], tuple(s for s in f["skus"].split(",") if s)) for f in filas
        ]
