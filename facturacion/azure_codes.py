"""Acceso a Azure SQL para consultar el catálogo de productos EFFI.

Este módulo centraliza la conexión con el repositorio de productos en Azure SQL,
realiza búsquedas por código de barras y texto libre, y expone utilidades para
que la lógica de importación y revisión de excepciones pueda validar productos
sin depender de EFFI directamente. La configuración se toma desde
facturacion/.env con variables como SQL_SERVER, SQL_DATABASE, SQL_USERNAME y
SQL_DRIVER.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
from collections import Counter
from contextlib import closing
from pathlib import Path
from typing import Any, Iterable

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))


def _connection_string() -> str:
    """Construye la cadena de conexión ODBC usando la configuración local del entorno."""
    server = os.getenv("SQL_SERVER", "").strip()
    database = os.getenv("SQL_DATABASE", "").strip()
    if not server or not database:
        raise RuntimeError("Configura SQL_SERVER y SQL_DATABASE en facturacion/.env.")

    driver = os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server").strip()
    encrypt = os.getenv("SQL_ENCRYPT", "yes").strip()
    trust_cert = os.getenv("SQL_TRUST_CERT", "no").strip()
    timeout = os.getenv("SQL_CONNECTION_TIMEOUT", "30").strip()

    parts = [
        f"DRIVER={{{driver}}};",
        f"SERVER={server};",
        f"DATABASE={database};",
        f"Encrypt={encrypt};",
        f"TrustServerCertificate={trust_cert};",
        f"Connection Timeout={timeout};",
    ]

    authentication = os.getenv("SQL_AUTHENTICATION", "").strip()
    username = os.getenv("SQL_USERNAME", "").strip()
    password = os.getenv("SQL_PASSWORD", "")

    if authentication:
        parts.append(f"Authentication={authentication};")
        if authentication == "ActiveDirectoryPassword":
            if not username or not password:
                raise RuntimeError(
                    "Para SQL_AUTHENTICATION=ActiveDirectoryPassword, configura "
                    "SQL_USERNAME y SQL_PASSWORD en facturacion/.env."
                )
            parts.extend((f"UID={username};", f"PWD={password};"))
    elif not username or not password:
        raise RuntimeError(
            "Configura SQL_USERNAME y SQL_PASSWORD, o SQL_AUTHENTICATION, "
            "en facturacion/.env."
        )
    else:
        parts.extend((f"UID={username};", f"PWD={password};"))

    return "".join(parts)


def _connect():
    import pyodbc

    connection_string = _connection_string()
    driver = os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server").strip()
    installed_drivers = pyodbc.drivers()
    if driver not in installed_drivers:
        if installed_drivers:
            available = ", ".join(installed_drivers)
            raise RuntimeError(
                f"ODBC driver {driver!r} is not installed. Available drivers: {available}. "
                "Set SQL_DRIVER to an installed driver."
            )
        raise RuntimeError(
            f"ODBC driver {driver!r} is not installed. Install unixODBC and the configured "
            "Microsoft ODBC driver, or set SQL_DRIVER to an installed driver."
        )
    return pyodbc.connect(connection_string)


def get_codes() -> list[str]:
    """Return non-empty barcodes currently stored in Azure SQL."""
    query = """
        SELECT CodigoBarras
        FROM [stg].[ProductoVerificador]
        WHERE CodigoBarras IS NOT NULL AND LTRIM(RTRIM(CodigoBarras)) <> ''
        ORDER BY CodigoBarras
    """
    with closing(_connect()) as connection:
        cursor = connection.cursor()
        cursor.execute(query)
        return [str(row[0]).strip() for row in cursor.fetchall()]


def get_product_by_code(code: str) -> dict[str, Any] | None:
    """Return the product fields used by the verifier, or None if not found."""
    query = """
        SELECT TOP 1
            IdProducto, CodigoBarras, Nombre, Referencia,
            Categoria, Marca, PrecioConImpuestoTarifaNormal
        FROM [stg].[ProductoVerificador]
        WHERE CodigoBarras = ?
    """
    with closing(_connect()) as connection:
        cursor = connection.cursor()
        cursor.execute(query, (code.strip(),))
        row = cursor.fetchone()
        if row is None:
            return None
        columns = [column[0] for column in cursor.description]
        product = dict(zip(columns, row))
        price = product.get("PrecioConImpuestoTarifaNormal")
        if price is not None:
            product["PrecioConImpuestoTarifaNormal"] = float(price)
        return product


def search_products(query_text: str, limit: int = 80) -> list[dict[str, Any]]:
    """Search Azure products by user-entered terms for exception review."""
    terms = list(dict.fromkeys(re.findall(r"[a-zA-Z0-9]+", (query_text or "").strip())))
    if not terms:
        return []

    terms = sorted(terms, key=len, reverse=True)[:8]
    fields = ("Nombre", "Referencia", "CodigoBarras", "Marca", "Categoria")
    term_conditions = [
        "(" + " OR ".join(f"{field} LIKE ?" for field in fields) + ")"
        for _ in terms
    ]
    order_conditions = " + ".join(
        f"CASE WHEN {condition} THEN 1 ELSE 0 END"
        for condition in term_conditions
    )
    query = f"""
        SELECT TOP {min(max(int(limit), 1), 100)}
            IdProducto, CodigoBarras, Nombre, Referencia, Categoria, Marca,
            PrecioConImpuestoTarifaNormal, PrecioConImpuestoTarifaMayorista
        FROM [stg].[ProductoVerificador]
        WHERE {' OR '.join(term_conditions)}
        ORDER BY {order_conditions} DESC, Nombre, IdProducto
    """
    patterns = [f"%{term}%" for term in terms for _ in fields]
    with closing(_connect()) as connection:
        with closing(connection.cursor()) as cursor:
            cursor.execute(query, patterns + patterns)
            columns = [column[0] for column in cursor.description]
            products: dict[Any, dict[str, Any]] = {}
            for row in cursor.fetchall():
                product = dict(zip(columns, row))
                if product.get("IdProducto") is not None:
                    products.setdefault(product["IdProducto"], product)
    return list(products.values())


def search_products_by_barcode(barcode: str) -> list[dict[str, Any]]:
    """Search only by an exact barcode match in the EFFI product catalog."""
    code = (barcode or "").strip()
    if not code:
        return []
    query = """
        SELECT
            IdProducto, CodigoBarras, Nombre, Referencia, Categoria, Marca,
            PrecioConImpuestoTarifaNormal, PrecioConImpuestoTarifaMayorista
        FROM [stg].[ProductoVerificador]
        WHERE CodigoBarras = ?
        ORDER BY IdProducto
    """
    with closing(_connect()) as connection:
        with closing(connection.cursor()) as cursor:
            cursor.execute(query, (code,))
            columns = [column[0] for column in cursor.description]
            return [dict(zip(columns, row)) for row in cursor.fetchall() if row[0] is not None]


def get_product_by_id(product_id: Any) -> dict[str, Any] | None:
    """Load one complete product record after a reviewer selects an ID."""
    query = """
        SELECT TOP 1
            IdProducto, CodigoBarras, Nombre, Referencia, Categoria, Marca,
            PrecioConImpuestoTarifaNormal, PrecioConImpuestoTarifaMayorista
        FROM [stg].[ProductoVerificador]
        WHERE IdProducto = ?
    """
    with closing(_connect()) as connection:
        with closing(connection.cursor()) as cursor:
            cursor.execute(query, (product_id,))
            row = cursor.fetchone()
            if row is None:
                return None
            columns = [column[0] for column in cursor.description]
            return dict(zip(columns, row))


def get_product_ids_by_barcodes(barcodes: Iterable[str]) -> dict[str, Any]:
    """Return EFFI product IDs indexed by matching barcode."""
    codes = list(dict.fromkeys(code.strip() for code in barcodes if code and code.strip()))
    if not codes:
        return {}

    product_ids_by_code: dict[str, set[Any]] = {}
    with closing(_connect()) as connection:
        with closing(connection.cursor()) as cursor:
            for offset in range(0, len(codes), 2000):
                batch = codes[offset : offset + 2000]
                placeholders = ", ".join("?" for _ in batch)
                query = f"""
                    SELECT CodigoBarras, IdProducto
                    FROM [stg].[ProductoVerificador]
                    WHERE CodigoBarras IN ({placeholders})
                """
                cursor.execute(query, batch)
                for barcode, product_id in cursor.fetchall():
                    if barcode is not None and product_id is not None:
                        product_ids_by_code.setdefault(str(barcode).strip(), set()).add(product_id)
    return {
        barcode: next(iter(product_ids))
        for barcode, product_ids in product_ids_by_code.items()
        if len(product_ids) == 1
    }


def _normalize_description(value: str) -> str:
    import unicodedata

    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(character for character in value if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def get_product_candidates_by_descriptions(
    descriptions: Iterable[str],
) -> dict[str, list[dict[str, Any]]]:
    """Return Azure products that share meaningful tokens with each description."""
    unique_descriptions = list(dict.fromkeys(
        str(value).strip() for value in descriptions if value and str(value).strip()
    ))
    normalized_descriptions = {
        _normalize_description(description): _normalize_description(description).split()
        for description in unique_descriptions
    }
    def is_searchable(token: str) -> bool:
        return (
            len(token) >= 3
            or (any(character.isalpha() for character in token) and any(character.isdigit() for character in token))
            or (token.isdigit() and len(token) >= 2)
        )

    token_frequency = Counter(
        token
        for description_tokens in normalized_descriptions.values()
        for token in set(description_tokens)
        if is_searchable(token)
    )
    selected_tokens = set()
    for description_tokens in normalized_descriptions.values():
        distinctive = sorted(
            {token for token in description_tokens if is_searchable(token)},
            key=lambda token: (token_frequency[token], -len(token), token),
        )
        selected_tokens.update(distinctive[:2])
    tokens = sorted(
        selected_tokens,
        key=lambda token: (token_frequency[token], -len(token), token),
    )[:160]
    if not tokens:
        return {}

    products: dict[Any, dict[str, Any]] = {}
    search_fields = ("Nombre", "Referencia", "Marca", "Categoria")
    token_batch_size = 400
    with closing(_connect()) as connection:
        with closing(connection.cursor()) as cursor:
            for offset in range(0, len(tokens), token_batch_size):
                batch = tokens[offset : offset + token_batch_size]
                term_conditions = [
                    "(" + " OR ".join(f"{field} LIKE ?" for field in search_fields) + ")"
                    for _ in batch
                ]
                query = f"""
                    SELECT
                        IdProducto, CodigoBarras, Nombre, Referencia, Categoria, Marca,
                        PrecioConImpuestoTarifaNormal, PrecioConImpuestoTarifaMayorista
                    FROM [stg].[ProductoVerificador]
                    WHERE IdProducto IS NOT NULL AND ({' OR '.join(term_conditions)})
                """
                patterns = [f"%{token}%" for token in batch for _ in search_fields]
                cursor.execute(query, patterns)
                columns = [column[0] for column in cursor.description]
                for row in cursor.fetchall():
                    product = dict(zip(columns, row))
                    products.setdefault(product["IdProducto"], product)

    products_with_tokens = [
        (
            product,
            set(_normalize_description(" ".join(
                str(product.get(field) or "") for field in ("Nombre", "Referencia", "Marca", "Categoria")
            )).split()),
        )
        for product in products.values()
    ]
    candidates_by_description: dict[str, list[dict[str, Any]]] = {}
    for normalized, description_tokens in normalized_descriptions.items():
        description_token_set = set(description_tokens)
        ranked = sorted(
            (
                (product, len(description_token_set & product_tokens))
                for product, product_tokens in products_with_tokens
                if description_token_set & product_tokens
            ),
            key=lambda item: item[1],
            reverse=True,
        )
        candidates_by_description[normalized] = [product for product, _ in ranked[:200]]

    return candidates_by_description


def get_product_by_description(description: str) -> dict[str, Any] | None:
    """Return a product only for one exact normalized name or reference match."""
    candidates = get_product_candidates_by_descriptions([description])
    normalized = _normalize_description(description)
    products = candidates.get(normalized, [])
    exact_matches = [
        product
        for product in products
        if any(
            _normalize_description(str(product.get(field) or "")) == normalized
            for field in ("Nombre", "Referencia")
        )
    ]
    return exact_matches[0] if len(exact_matches) == 1 else None


def get_product_ids_by_descriptions(
    descriptions: Iterable[str],
) -> dict[str, list[dict[str, Any]]]:
    """Return full product candidates indexed by normalized description."""
    candidates = get_product_candidates_by_descriptions(descriptions)
    return {
        normalized: [product for product in products if product.get("IdProducto") is not None]
        for normalized, products in candidates.items()
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Consulta códigos de barras en Azure SQL.")
    parser.add_argument("--code", help="Busca un producto por su código de barras.")
    parser.add_argument("--output", type=Path, help="Guarda todos los códigos en un CSV.")
    args = parser.parse_args()

    if args.code:
        product = get_product_by_code(args.code)
        if product is None:
            print("Código no encontrado.")
            return 1
        for field, value in product.items():
            print(f"{field}: {value}")
        return 0

    codes = get_codes()
    if args.output:
        with args.output.open("w", newline="", encoding="utf-8-sig") as output:
            writer = csv.writer(output)
            writer.writerow(["CodigoBarras"])
            writer.writerows((code,) for code in codes)
        print(f"{len(codes)} códigos guardados en {args.output}")
    else:
        for code in codes:
            print(code)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())