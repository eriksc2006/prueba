from __future__ import annotations

import argparse
import csv
import os
from contextlib import closing
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))


def _connection_string() -> str:
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
    if authentication:
        parts.append(f"Authentication={authentication};")
    else:
        username = os.getenv("SQL_USERNAME", "").strip()
        password = os.getenv("SQL_PASSWORD", "")
        if not username or not password:
            raise RuntimeError(
                "Configura SQL_USERNAME y SQL_PASSWORD, o SQL_AUTHENTICATION, "
                "en facturacion/.env."
            )
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