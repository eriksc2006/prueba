"""Generación de libros de importación EFFI desde XML UBL/DIAN.

Este módulo procesa facturas de proveedores en formato DIAN/UBL, normaliza el
detalle de líneas y crea el libro Excel usado como plantilla de importación en
EFFI. La implementación cubre los casos más comunes del negocio colombiano:

- extracción de Invoice o AttachedDocument
- reconocimiento de códigos de barra, item del comprador y código de proveedor
- reconstrucción de precio/descuento/neto cuando el XML trae bonificaciones o
  ajustes de redondeo
- consolidación de líneas repetidas por producto para preservar el neto total
  y calcular costo efectivo
- mapeo de tasas tributarias a códigos EFFI configurables
- generación del Excel con advertencias y excepciones para revisión humana

La intención es mantener el XML como fuente de verdad para el valor neto y, al
mismo tiempo, producir un archivo listo para revisión/importación manual en EFFI.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Mapping

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


ROOT = Path(__file__).resolve().parent
DEFAULT_XML_DIR = ROOT / "files" / "xml"
DEFAULT_TEMPLATE = ROOT / "files" / "template" / "plantilla_importacion_conceptos_compra.xlsx"
DEFAULT_OUTPUT_DIR = ROOT / "output"
MAX_XML_BYTES = 20 * 1024 * 1024
CENT = Decimal("0.02")
PRICE_RECONCILIATION_LIMIT = Decimal("0.10")

# Code 1 is the 19% tax code shown in the supplied EFFI template example.
# Confirm this identifier against the tax setup in the target EFFI company.
DEFAULT_TAX_CODES = {Decimal("19"): 1}


class InvoiceParseError(ValueError):
    """Raised when an input cannot be interpreted as a DIAN invoice."""


@dataclass(frozen=True)
class InvoiceLine:
    """Representa una línea de factura DIAN con su información financiera y fiscal.

    Se usa como estructura central para el cálculo del neto, el desglosamiento de
    descuentos, la identificación del producto y la generación del Excel de
    importación. Los valores se mantienen en Decimals para evitar errores de
    redondeo en montos de dinero.
    """

    line_id: str
    item_code: str
    barcode: str
    description: str
    quantity: Decimal
    unit: str
    unit_price: Decimal
    discount_total: Decimal
    net_amount: Decimal
    tax_rates: tuple[Decimal, ...]


@dataclass(frozen=True)
class Invoice:
    """Cabecera de factura con todas sus líneas de detalle."""

    invoice_id: str
    issue_date: str
    supplier_name: str
    supplier_tax_id: str
    currency: str
    lines: tuple[InvoiceLine, ...]


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(element: ET.Element, name: str) -> list[ET.Element]:
    return [child for child in element if _local_name(child.tag) == name]


def _child(element: ET.Element | None, name: str) -> ET.Element | None:
    if element is None:
        return None
    return next((child for child in element if _local_name(child.tag) == name), None)


def _path(element: ET.Element | None, *names: str) -> ET.Element | None:
    for name in names:
        element = _child(element, name)
        if element is None:
            return None
    return element


def _text(element: ET.Element | None) -> str:
    return (element.text or "").strip() if element is not None else ""


def _value(element: ET.Element | None, *names: str) -> str:
    return _text(_path(element, *names))


def _decimal(value: str, label: str) -> Decimal:
    try:
        result = Decimal(value)
    except (InvalidOperation, ValueError) as exc:
        raise InvoiceParseError(f"Valor numérico inválido para {label}: {value!r}") from exc
    if not result.is_finite():
        raise InvoiceParseError(f"Valor no finito para {label}: {value!r}")
    return result


def _is_invoice(element: ET.Element) -> bool:
    return _local_name(element.tag) == "Invoice"


def _find_invoice(root: ET.Element) -> ET.Element:
    if _is_invoice(root):
        return root

    # DIAN AttachedDocument puts the UBL Invoice in a cbc:Description CDATA
    # block. Other Description blocks can contain acknowledgements, so each
    # candidate is parsed and checked by its root element.
    for element in root.iter():
        if _local_name(element.tag) != "Description":
            continue
        payload = _text(element)
        if "<Invoice" not in payload:
            continue
        start = payload.find("<Invoice")
        end = payload.rfind("</Invoice>")
        if end < start:
            continue
        candidate = payload[start : end + len("</Invoice>")]
        try:
            invoice = ET.fromstring(candidate)
        except ET.ParseError:
            continue
        if _is_invoice(invoice):
            return invoice

    # Accommodate wrappers that contain Invoice as a regular nested element.
    for element in root.iter():
        if _is_invoice(element):
            return element
    raise InvoiceParseError("No se encontró un elemento UBL Invoice en el XML.")


def _item_code(item: ET.Element | None) -> str:
    if item is None:
        return ""
    # The buyer-assigned code is the best chance of being the internal EFFI
    # article ID. Otherwise use the standard ID (often the GTIN), then the
    # seller's catalog code. Ignore common empty/placeholder values.
    for identifier_type in (
        "BuyersItemIdentification",
        "StandardItemIdentification",
        "SellersItemIdentification",
    ):
        code = _value(_child(item, identifier_type), "ID")
        if code and code not in {"0", "-"}:
            return code
    return ""


def _barcode(item: ET.Element | None) -> str:
    if item is None:
        return ""
    identifier = _child(item, "StandardItemIdentification")
    if identifier is None:
        return ""
    code_element = _child(identifier, "ID")
    if code_element is None:
        return ""
    scheme_id = code_element.attrib.get("schemeID", "").strip()
    scheme_name = code_element.attrib.get("schemeName", "").strip().casefold()
    if scheme_id != "010" and scheme_name != "gtin":
        return ""
    code = _text(code_element)
    if not code.isdigit() or len(code) not in {8, 12, 13, 14}:
        return ""
    return code


def _line_tax_rates(line: ET.Element) -> tuple[Decimal, ...]:
    rates: set[Decimal] = set()
    for tax_total in _children(line, "TaxTotal"):
        for subtotal in _children(tax_total, "TaxSubtotal"):
            percent = _value(subtotal, "TaxCategory", "Percent")
            if percent:
                rates.add(_decimal(percent, "porcentaje de impuesto"))
    return tuple(sorted(rates))


def _parse_line(element: ET.Element) -> InvoiceLine:
    quantity_element = _child(element, "InvoicedQuantity")
    quantity_text = _text(quantity_element)
    if not quantity_text:
        raise InvoiceParseError(f"La línea {_value(element, 'ID') or '?'} no tiene cantidad.")
    quantity = _decimal(quantity_text, "cantidad")

    net_text = _value(element, "LineExtensionAmount")
    if not net_text:
        raise InvoiceParseError(f"La línea {_value(element, 'ID') or '?'} no tiene valor neto.")
    net_amount = _decimal(net_text, "valor neto de línea")

    price_text = _value(element, "Price", "PriceAmount")
    unit_price = _decimal(price_text, "precio unitario") if price_text else (
        net_amount / quantity if quantity else Decimal("0")
    )

    discount_total = Decimal("0")
    for allowance in _children(element, "AllowanceCharge"):
        indicator = _value(allowance, "ChargeIndicator").lower()
        # EFFI exposes a discount field, not a line surcharge field.
        if indicator == "false":
            amount = _value(allowance, "Amount")
            if amount:
                discount_total += _decimal(amount, "descuento")

    # Some suppliers mark promotional/free quantities with a zero PriceAmount
    # but include the full list value as a 100% AllowanceCharge. Recover that
    # gross unit price so EFFI receives a positive price and matching discount.
    if unit_price == 0 and quantity > 0 and discount_total > 0 and net_amount == 0:
        unit_price = discount_total / quantity

    # Some suppliers omit AllowanceCharge but encode the line discount as the
    # difference between unit price × quantity and LineExtensionAmount.
    if discount_total == 0 and quantity > 0:
        implied_discount = unit_price * quantity - net_amount
        if implied_discount > CENT:
            discount_total = implied_discount

    # Reconcile small UBL rounding differences so the imported unit price and
    # discount reproduce the authoritative line net amount. Larger differences
    # are retained for manual review instead of silently changing the price.
    if quantity > 0:
        expected_net = unit_price * quantity - discount_total
        if CENT < abs(expected_net - net_amount) <= PRICE_RECONCILIATION_LIMIT:
            unit_price = (net_amount + discount_total) / quantity

    rates = _line_tax_rates(element)
    if not rates:
        tax_amounts = [
            _value(total, "TaxAmount")
            for total in _children(element, "TaxTotal")
            if _value(total, "TaxAmount")
        ]
        if tax_amounts and all(_decimal(amount, "impuesto") == 0 for amount in tax_amounts):
            rates = (Decimal("0"),)

    item = _child(element, "Item")
    return InvoiceLine(
        line_id=_value(element, "ID"),
        item_code=_item_code(item),
        barcode=_barcode(item),
        description=_value(item, "Description"),
        quantity=quantity,
        unit=quantity_element.attrib.get("unitCode", ""),
        unit_price=unit_price,
        discount_total=discount_total,
        net_amount=net_amount,
        tax_rates=rates,
    )


def parse_invoice(path: Path) -> Invoice:
    """Read an Invoice or DIAN AttachedDocument containing one Invoice."""
    size = path.stat().st_size
    if size > MAX_XML_BYTES:
        raise InvoiceParseError(f"El archivo supera el límite de {MAX_XML_BYTES // (1024 * 1024)} MB.")
    try:
        root = ET.fromstring(path.read_bytes())
    except (ET.ParseError, OSError) as exc:
        raise InvoiceParseError(f"No se pudo leer el XML: {exc}") from exc

    element = _find_invoice(root)
    lines = tuple(_parse_line(line) for line in _children(element, "InvoiceLine"))
    if not lines:
        raise InvoiceParseError("La factura no contiene líneas de detalle.")

    supplier_party = _path(element, "AccountingSupplierParty", "Party")
    supplier_name = _value(supplier_party, "PartyName", "Name")
    if not supplier_name:
        supplier_name = _value(supplier_party, "PartyTaxScheme", "RegistrationName")
    supplier_tax_id = _value(supplier_party, "PartyTaxScheme", "CompanyID")
    if not supplier_tax_id:
        supplier_tax_id = _value(supplier_party, "PartyIdentification", "ID")

    return Invoice(
        invoice_id=_value(element, "ID"),
        issue_date=_value(element, "IssueDate"),
        supplier_name=supplier_name,
        supplier_tax_id=supplier_tax_id,
        currency=_value(element, "DocumentCurrencyCode"),
        lines=lines,
    )


def _tax_code(line: InvoiceLine, mapping: dict[Decimal, int]) -> int | None:
    if len(line.tax_rates) != 1:
        return None
    normalized_rate = line.tax_rates[0].normalize()
    return next((code for rate, code in mapping.items() if rate.normalize() == normalized_rate), None)


def _normalize_description(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(character for character in value if not unicodedata.combining(character))
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def _presentation_tokens(value: str) -> set[str]:
    return {
        re.sub(r"\s+", "", match.group(0).lower()).replace(",", ".")
        for match in re.finditer(
            r"\b(?:px\s*\d+|\d+(?:[.,]\d+)?\s*(?:ml|lt|l|gr|g|kg|mg|oz|und|unidades?))\b",
            value.lower(),
        )
    }


def _candidate_score(line: InvoiceLine, product: Mapping[str, Any]) -> float:
    description = _normalize_description(line.description)
    name = _normalize_description(str(product.get("Nombre") or ""))
    reference = _normalize_description(str(product.get("Referencia") or ""))
    brand = _normalize_description(str(product.get("Marca") or ""))
    category = _normalize_description(str(product.get("Categoria") or ""))
    source_tokens = set(description.split())
    name_tokens = set((name + " " + reference).split())
    catalog_tokens = set((name + " " + reference + " " + brand + " " + category).split())
    if not source_tokens or not catalog_tokens:
        return 0.0

    overlap = len(source_tokens & catalog_tokens) / len(source_tokens)
    jaccard = len(source_tokens & name_tokens) / max(len(source_tokens | name_tokens), 1)
    text_similarity = SequenceMatcher(None, description, name + " " + reference).ratio()
    brand_tokens = set(brand.split())
    brand_match = len(brand_tokens & source_tokens) / len(brand_tokens) if brand_tokens else None

    source_presentation = _presentation_tokens(line.description)
    product_presentation = _presentation_tokens(str(product.get("Nombre") or ""))
    if not product_presentation:
        product_presentation = _presentation_tokens(str(product.get("Referencia") or ""))
    presentation_match = (
        float(bool(source_presentation & product_presentation))
        if source_presentation and product_presentation
        else None
    )

    invoice_price = line.unit_price
    if len(line.tax_rates) == 1:
        invoice_price *= Decimal("1") + line.tax_rates[0] / Decimal("100")
    catalog_prices = [
        Decimal(str(product[field]))
        for field in ("PrecioConImpuestoTarifaNormal", "PrecioConImpuestoTarifaMayorista")
        if product.get(field) is not None
    ]
    price_match = None
    if catalog_prices and invoice_price > 0:
        relative_difference = min(
            abs(price - invoice_price) / max(abs(invoice_price), Decimal("1"))
            for price in catalog_prices
        )
        price_match = max(0.0, 1.0 - float(relative_difference))

    components = [(overlap, 0.42), (jaccard, 0.08), (text_similarity, 0.14)]
    if brand_match is not None:
        components.append((brand_match, 0.12))
    if presentation_match is not None:
        components.append((presentation_match, 0.14))
    if price_match is not None:
        components.append((price_match, 0.10))
    total_weight = sum(weight for _, weight in components)
    return sum(value * weight for value, weight in components) / total_weight


def _variant_tokens(value: str) -> set[str]:
    variants = re.findall(r"(?<!\d)\d{1,3}\s*[.,/-]\s*\d{1,3}[a-z]?(?!\d)", value or "", re.IGNORECASE)
    return {
        re.sub(r"\s+", "", variant).replace(",", ".").replace("/", ".").replace("-", ".").lower()
        for variant in variants
    }


def _supplier_matches_product_brand(supplier_name: str, brand: str) -> bool:
    legal_suffixes = {"s", "a", "sas", "ltda", "limitada", "inc", "llc", "corp", "corporation"}
    supplier_tokens = [
        token for token in _normalize_description(supplier_name).split()
        if token not in legal_suffixes
    ]
    brand_tokens = _normalize_description(brand).split()
    return bool(supplier_tokens and brand_tokens and "".join(supplier_tokens) == "".join(brand_tokens))


def _confident_no_barcode_match(
    line: InvoiceLine,
    product: Mapping[str, Any],
    margin: float,
    supplier_name: str = "",
) -> bool:
    if line.barcode or product.get("IdProducto") is None or margin < 0.10:
        return False
    source_variants = _variant_tokens(line.description)
    product_text = " ".join((str(product.get("Nombre") or ""), str(product.get("Referencia") or "")))
    if not source_variants or not source_variants.intersection(_variant_tokens(product_text)):
        return False
    source_tokens = set(_normalize_description(line.description).split())
    product_tokens = set(_normalize_description(product_text).split())
    if not any(len(token) >= 5 for token in source_tokens & product_tokens):
        return False
    brand = str(product.get("Marca") or "")
    return bool(brand and _supplier_matches_product_brand(supplier_name, brand))


def _description_match(
    line: InvoiceLine,
    product_ids_by_description: Mapping[str, Any] | None,
) -> tuple[Mapping[str, Any] | None, float, float]:
    if not line.description or product_ids_by_description is None:
        return None, 0.0, 0.0
    value = product_ids_by_description.get(_normalize_description(line.description))
    if value is None:
        return None, 0.0, 0.0
    if not isinstance(value, (list, tuple)):
        if isinstance(value, Mapping):
            value = [value]
        else:
            return None, 0.0, 0.0

    ranked = sorted(
        (
            (candidate, _candidate_score(line, candidate))
            for candidate in value
            if isinstance(candidate, Mapping)
        ),
        key=lambda item: item[1],
        reverse=True,
    )
    if not ranked:
        return None, 0.0, 0.0
    best, score = ranked[0]
    runner_up_score = ranked[1][1] if len(ranked) > 1 else 0.0
    return best, score, score - runner_up_score


def _resolved_product(
    line: InvoiceLine,
    product_ids_by_barcode: Mapping[str, Any] | None,
    product_ids_by_description: Mapping[str, Any] | None,
    supplier_name: str = "",
) -> tuple[Any, Mapping[str, Any] | None, float, float, bool]:
    if line.barcode:
        if product_ids_by_barcode is None:
            return None, None, 0.0, 0.0, False
        barcode_product_id = product_ids_by_barcode.get(line.barcode)
        if barcode_product_id is not None:
            return barcode_product_id, None, 1.0, 1.0, True
        return None, None, 0.0, 0.0, False
    candidate, score, margin = _description_match(line, product_ids_by_description)
    accepted = bool(
        candidate
        and candidate.get("IdProducto") is not None
        and (
            (score >= 0.72 and (margin >= 0.06 or margin == 1.0))
            or _confident_no_barcode_match(line, candidate, margin, supplier_name)
        )
    )
    if accepted:
        return candidate["IdProducto"], candidate, score, margin, True
    return None, candidate, score, margin, False


def _import_amounts(line: InvoiceLine) -> tuple[Decimal, Decimal, Decimal]:
    """Keep the XML line net amount authoritative in the values sent to EFFI."""
    price = line.unit_price
    discount = line.discount_total
    source_delta = price * line.quantity - discount - line.net_amount
    if source_delta and line.quantity > 0:
        adjusted_discount = price * line.quantity - line.net_amount
        if adjusted_discount >= 0:
            discount = adjusted_discount
        else:
            price = (line.net_amount + discount) / line.quantity
    return price, discount, source_delta


def _is_full_discount_bonus(line: InvoiceLine) -> bool:
    gross_amount = line.unit_price * line.quantity
    return (
        line.quantity > 0
        and line.net_amount == 0
        and gross_amount > 0
        and line.discount_total > 0
        and abs(gross_amount - line.discount_total) <= CENT
    )


def _group_import_lines(
    invoice: Invoice,
    product_ids_by_barcode: Mapping[str, Any] | None,
    product_ids_by_description: Mapping[str, Any] | None,
) -> list[tuple[InvoiceLine, ...]]:
    """Agrupa líneas por producto para exportar un único registro a EFFI.

    La agrupación evita duplicar artículos cuando una misma referencia aparece en
    varias líneas de factura. Cuando el sistema puede verificar que todas las
    líneas pertenecen al mismo producto y comparten el perfil fiscal, se calcula
    un costo efectivo combinado; en caso contrario se conservan separadas para no
    distorsionar el neto de la factura.
    """
    grouped: dict[tuple[str, str, str], list[InvoiceLine]] = {}
    for line in invoice.lines:
        if line.barcode:
            identity = ("barcode", line.barcode)
        else:
            product_id, _, _, _, resolved = _resolved_product(
                line, product_ids_by_barcode, product_ids_by_description
            )
            if resolved and product_id is not None:
                identity = ("effi", str(product_id))
            elif line.item_code and line.item_code not in {"0", "-"}:
                identity = ("item", line.item_code.strip().casefold())
            else:
                # Never reuse the supplier's item reference as an EFFI article ID.
                # It may group matching invoice lines, but is never written to the
                # EFFI article column unless Azure verifies the product.
                description = _normalize_description(line.description)
                identity = ("description", description or line.line_id)
        key = (*identity, line.unit.strip().casefold())
        grouped.setdefault(key, []).append(line)

    result: list[tuple[InvoiceLine, ...]] = []
    for lines in grouped.values():
        charged_tax_profiles = {
            tuple(rate.normalize() for rate in line.tax_rates)
            for line in lines
            if line.net_amount != 0
        }
        can_allocate_bonus = (
            len(lines) > 1
            and all(line.quantity > 0 and line.net_amount >= 0 for line in lines)
            and sum((line.net_amount for line in lines), Decimal("0")) > 0
            and len(charged_tax_profiles) <= 1
        )
        if can_allocate_bonus:
            result.append(tuple(lines))
        else:
            result.extend((line,) for line in lines)
    return result


def _import_group_amounts(lines: tuple[InvoiceLine, ...]) -> tuple[Decimal, Decimal, Decimal, InvoiceLine]:
    tax_line = next((line for line in lines if line.net_amount != 0), lines[0])
    if len(lines) == 1:
        price, discount, source_delta = _import_amounts(lines[0])
        return lines[0].quantity, price, discount, tax_line
    quantity = sum((line.quantity for line in lines), Decimal("0"))
    net_amount = sum((line.net_amount for line in lines), Decimal("0"))
    return quantity, net_amount / quantity, Decimal("0"), tax_line


def _line_warnings(
    line: InvoiceLine,
    mapping: dict[Decimal, int],
    product_ids_by_barcode: Mapping[str, Any] | None = None,
    product_ids_by_description: Mapping[str, Any] | None = None,
    supplier_name: str = "",
) -> list[str]:
    warnings: list[str] = []
    label = f"línea {line.line_id or '?'}"
    if not line.item_code:
        warnings.append(f"{label}: no se encontró código de artículo; completar en EFFI.")
    if product_ids_by_barcode is not None or product_ids_by_description is not None:
        _, _, score, margin, resolved = _resolved_product(
            line, product_ids_by_barcode, product_ids_by_description, supplier_name
        )
        if not resolved:
            reason = (
                f"Código de barras {line.barcode} no registrado en EFFI para el producto "
                f"'{line.description or 'sin descripción'}'; vincular este código en EFFI"
                if line.barcode and product_ids_by_barcode is not None
                else "descripción sin coincidencia confiable"
            )
            if score:
                reason += f" (candidato más cercano: {score:.0%}, diferencia {margin:.0%})"
            warnings.append(f"{label}: {reason}; revisar artículo.")
    if line.quantity <= 0:
        warnings.append(f"{label}: cantidad no positiva ({line.quantity}).")
    if len(line.tax_rates) != 1:
        warnings.append(f"{label}: tasa de impuesto ausente o múltiple; revisar código de impuesto.")
    elif _tax_code(line, mapping) is None:
        warnings.append(f"{label}: tasa {line.tax_rates[0]}% sin código EFFI configurado.")
    import_price, import_discount, source_delta = _import_amounts(line)
    expected_net = import_price * line.quantity - import_discount
    if abs(expected_net - line.net_amount) > Decimal("0.005"):
        warnings.append(
            f"{label}: no se pudo conciliar el neto importado ({expected_net}) con el neto XML "
            f"({line.net_amount}); revisar la línea antes de importarla."
        )
    elif source_delta:
        warnings.append(f"{label}: ajuste de {source_delta} para igualar el neto declarado en el XML.")
    return warnings


def _write_detail(
    sheet: Worksheet,
    invoice: Invoice,
    tax_codes: dict[Decimal, int],
    product_ids_by_barcode: Mapping[str, Any] | None = None,
    product_ids_by_description: Mapping[str, Any] | None = None,
) -> tuple[dict[str, int], list[tuple[InvoiceLine, ...]]]:
    """Escribe el detalle de la factura sobre la hoja de importación EFFI.

    La plantilla del libro se reutiliza para conservar formato y columnas, pero
    se limpia el contenido de filas de ejemplo para poblar únicamente las líneas
    reales de la factura. Cada fila representa un producto o un grupo de líneas
    compatibles, y se persiste el precio, descuento y código tributario que
    reproduzcan el neto XML sin alterar la factura original.
    """
    # The supplied workbook has sample lines in rows 2-4. Remove cell values
    # only; preserve the template's formatting and its header row.
    for row in sheet.iter_rows(min_row=2, max_row=max(sheet.max_row, len(invoice.lines) + 1), min_col=1, max_col=9):
        for cell in row:
            cell.value = None

    style_source = [sheet.cell(row=2, column=column) for column in range(1, 10)]
    for row_number in range(2, max(sheet.max_row, len(invoice.lines) + 1) + 1):
        for column in range(2, 6):
            sheet.cell(row=row_number, column=column, value=None)
    grouped_lines = _group_import_lines(invoice, product_ids_by_barcode, product_ids_by_description)
    import_row_by_line: dict[str, int] = {}
    for row_number, source_lines in enumerate(grouped_lines, start=2):
        line = source_lines[0]
        resolved_id, _, _, _, _ = _resolved_product(
            line, product_ids_by_barcode, product_ids_by_description, invoice.supplier_name
        )
        quantity, import_price, import_discount, tax_line = _import_group_amounts(source_lines)
        for source_line in source_lines:
            import_row_by_line[source_line.line_id] = row_number
        is_barcode_mapped = (
            bool(line.barcode)
            and product_ids_by_barcode is not None
            and line.barcode in product_ids_by_barcode
        )
        values = (
            resolved_id,                # A: EFFI article ID
            None if is_barcode_mapped else 1,  # B: default exit type; leave blank for Azure barcode matches
            None,                       # C: lot (not in DIAN UBL)
            None,                       # D: serial (not in DIAN UBL)
            None,                       # E: observation (kept for manual entry)
            quantity,                   # F: quantity
            import_price,               # G: unit price
            import_discount,            # H: total line discount
            _tax_code(tax_line, tax_codes), # I: configured EFFI tax ID
        )
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_number, column=column, value=value)
            if row_number > 2:
                cell._style = style_source[column - 1]._style
                if style_source[column - 1].number_format:
                    cell.number_format = style_source[column - 1].number_format
    return import_row_by_line, grouped_lines


def _json_invoice(invoice: Invoice) -> dict[str, object]:
    result = asdict(invoice)
    for line in result["lines"]:
        for key in ("quantity", "unit_price", "discount_total", "net_amount"):
            line[key] = str(line[key])
        line["tax_rates"] = [str(rate) for rate in line["tax_rates"]]
    return result


def _parse_tax_mapping(values: Iterable[str]) -> dict[Decimal, int]:
    mapping = dict(DEFAULT_TAX_CODES)
    for value in values:
        try:
            rate_text, code_text = value.split("=", maxsplit=1)
            rate = Decimal(rate_text.strip())
            code = int(code_text.strip())
        except (ValueError, InvalidOperation) as exc:
            raise argparse.ArgumentTypeError(
                f"Mapeo inválido {value!r}; use el formato TASA=CODIGO, por ejemplo 19=1."
            ) from exc
        mapping[rate] = code
    return mapping


def _safe_stem(source: Path, invoice_id: str) -> str:
    base = invoice_id or source.stem
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("._")
    return safe or source.stem


def _invoice_output_stem(source: Path, invoice_id: str, supplier_tax_id: str = "") -> str:
    base = _safe_stem(source, invoice_id)
    origin = _safe_stem(source, supplier_tax_id or source.stem)
    return f"{base}_{origin}" if origin != base else base


def generate_workbook(
    source: Path,
    template: Path,
    output: Path,
    tax_codes: dict[Decimal, int],
    write_json: bool = False,
    product_ids_by_barcode: Mapping[str, Any] | None = None,
    product_ids_by_description: Mapping[str, Any] | None = None,
) -> list[str]:
    invoice = parse_invoice(source)
    workbook = load_workbook(template)
    sheet = workbook.active
    import_row_by_line, grouped_lines = _write_detail(
        sheet, invoice, tax_codes, product_ids_by_barcode, product_ids_by_description
    )
    group_by_line = {
        line.line_id: lines
        for lines in grouped_lines
        for line in lines
    }
    barcode_counts: dict[str, int] = {}
    for line in invoice.lines:
        if line.barcode:
            barcode_counts[line.barcode] = barcode_counts.get(line.barcode, 0) + 1
    unresolved_lookup = product_ids_by_barcode is not None or product_ids_by_description is not None
    exceptions = []
    automatic_matches = []
    for line in invoice.lines:
        _, candidate, score, margin, resolved = _resolved_product(
            line, product_ids_by_barcode, product_ids_by_description, invoice.supplier_name
        )
        if not line.barcode and candidate and resolved:
            automatic_matches.append((line, candidate, score, margin))
        source_group = group_by_line[line.line_id]
        group_quantity, group_price, _, _ = _import_group_amounts(source_group)
        grouped_bonus = len(source_group) > 1
        repeated_barcode = bool(line.barcode and barcode_counts.get(line.barcode, 0) > 1)
        full_discount_bonus = _is_full_discount_bonus(line)
        import_price, import_discount, source_delta = _import_amounts(line)
        if (unresolved_lookup and not resolved) or source_delta or line.quantity <= 0 or grouped_bonus or full_discount_bonus or repeated_barcode:
            reasons = []
            if unresolved_lookup and not resolved:
                if line.barcode:
                    reasons.append(
                        f"Código de barras no registrado en EFFI para el producto "
                        f"'{line.description or 'sin descripción'}'; vincular este código en EFFI"
                    )
                else:
                    reasons.append("Sin coincidencia confiable" if not score else "Candidato cercano; validar identidad")
            if source_delta:
                reasons.append("Neto XML conciliado; revisar diferencia fuente")
            if line.quantity <= 0:
                reasons.append("Cantidad no positiva")
            if grouped_bonus:
                reasons.append(
                    f"Bonificación/prorrateo: {group_quantity} unidades totales; "
                    f"costo unitario efectivo {group_price}"
                )
            if full_discount_bonus:
                reasons.append("Bonificación 100% detectada; vincular al producto facturado correspondiente")
            if repeated_barcode and not grouped_bonus:
                reasons.append("GTIN repetido con perfiles o unidades separadas; revisar vínculo y bonificación")
            exceptions.append((line, candidate, score, margin, resolved, source_delta, import_discount, "; ".join(reasons)))
    if exceptions:
        exception_sheet = workbook.create_sheet("Excepciones")
        exception_sheet.append([
            "Línea", "Tipo", "Código de barras", "Descripción factura", "Cantidad",
            "Presentación factura", "Precio unitario XML", "Descuento XML", "Neto XML", "Neto calculado XML",
            "Ajuste aplicado al descuento", "IdProducto candidato", "Nombre candidato",
            "Código de barras candidato", "Presentación candidato", "Referencia", "Marca", "Categoría",
            "Precio normal", "Precio mayorista",
            "Puntaje", "Decisión", "Fila importación", "Unidades totales del producto",
            "Neto agrupado", "Costo unitario efectivo", "Bonificación consolidada",
            "Líneas XML consolidadas",
        ])
        for line, candidate, score, margin, resolved, source_delta, import_discount, reason in exceptions:
            nearest = candidate or {}
            source_group = group_by_line[line.line_id]
            group_quantity, group_price, _, _ = _import_group_amounts(source_group)
            exception_sheet.append([
                line.line_id,
                reason,
                line.barcode,
                line.description,
                float(line.quantity),
                ", ".join(sorted(_presentation_tokens(line.description))),
                float(line.unit_price),
                float(line.discount_total),
                float(line.net_amount),
                float(line.unit_price * line.quantity - line.discount_total),
                float(import_discount - line.discount_total),
                nearest.get("IdProducto"),
                nearest.get("Nombre"),
                nearest.get("CodigoBarras"),
                ", ".join(sorted(_presentation_tokens(
                    str(nearest.get("Nombre") or "") + " " + str(nearest.get("Referencia") or "")
                ))),
                nearest.get("Referencia"),
                nearest.get("Marca"),
                nearest.get("Categoria"),
                nearest.get("PrecioConImpuestoTarifaNormal"),
                nearest.get("PrecioConImpuestoTarifaMayorista"),
                score,
                "Usado" if resolved else f"Revisar; margen {margin:.0%}",
                import_row_by_line[line.line_id],
                group_quantity,
                sum((source_line.net_amount for source_line in group_by_line[line.line_id]), Decimal("0")),
                group_price,
                len(group_by_line[line.line_id]) > 1,
                ", ".join(source_line.line_id for source_line in group_by_line[line.line_id]),
            ])
    if automatic_matches:
        match_sheet = workbook.create_sheet("Coincidencias automáticas")
        match_sheet.append([
            "Línea XML", "Descripción factura", "IdProducto EFFI", "Producto EFFI",
            "Código de barras EFFI", "Referencia EFFI", "Marca EFFI", "Categoría EFFI",
            "Puntaje", "Margen",
        ])
        for line, candidate, score, margin in automatic_matches:
            match_sheet.append([
                line.line_id,
                line.description,
                candidate.get("IdProducto"),
                candidate.get("Nombre"),
                candidate.get("CodigoBarras"),
                candidate.get("Referencia"),
                candidate.get("Marca"),
                candidate.get("Categoria"),
                score,
                margin,
            ])
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    workbook.close()

    if write_json:
        output.with_suffix(".json").write_text(
            json.dumps(_json_invoice(invoice), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    warnings = [
        warning
        for line in invoice.lines
        for warning in _line_warnings(
            line, tax_codes, product_ids_by_barcode, product_ids_by_description, invoice.supplier_name
        )
    ]
    warnings.extend(
        f"Producto {lines[0].description or lines[0].item_code}: Bonificación consolidada; "
        f"{_import_group_amounts(lines)[0]} unidades, costo efectivo {_import_group_amounts(lines)[1]} por unidad."
        for lines in grouped_lines
        if len(lines) > 1
    )
    return warnings


def _sources(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() == ".xml" else []
    if input_path.is_dir():
        return sorted(path for path in input_path.glob("*.xml") if path.is_file())
    return []


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Convierte líneas de XML UBL/DIAN a la plantilla de detalle de compras de EFFI."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_XML_DIR, help="XML individual o carpeta de XML.")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="Plantilla Excel de EFFI.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Carpeta destino.")
    parser.add_argument(
        "--tax-code",
        action="append",
        default=[],
        metavar="TASA=CODIGO",
        help="Mapeo adicional/sobrescrito de tasa DIAN a código EFFI (repetible). Ej.: 5=3.",
    )
    parser.add_argument("--json", action="store_true", help="Guarda también el detalle estructurado como JSON.")
    parser.add_argument(
        "--azure-lookup",
        action="store_true",
        help="Busca los códigos de barras en Azure SQL y escribe el IdProducto EFFI encontrado.",
    )
    args = parser.parse_args(argv)

    if not args.template.is_file():
        parser.error(f"No existe la plantilla: {args.template}")
    sources = _sources(args.input)
    if not sources:
        parser.error(f"No se encontraron XML en: {args.input}")
    try:
        tax_codes = _parse_tax_mapping(args.tax_code)
    except argparse.ArgumentTypeError as exc:
        parser.error(str(exc))

    failures = 0
    invoices: dict[Path, Invoice] = {}
    product_ids_by_barcode: Mapping[str, Any] | None = None
    product_ids_by_description: Mapping[str, Any] | None = None
    if args.azure_lookup:
        for source in sources:
            try:
                invoices[source] = parse_invoice(source)
            except (InvoiceParseError, OSError, ValueError) as exc:
                failures += 1
                print(f"ERROR {source.name}: {exc}", file=sys.stderr)
        try:
            from .azure_codes import get_product_candidates_by_descriptions, get_product_ids_by_barcodes

            product_ids_by_barcode = get_product_ids_by_barcodes(
                line.barcode for invoice in invoices.values() for line in invoice.lines if line.barcode
            )
            # PRUEBAS: búsqueda por descripción habilitada temporalmente para validar
            # coincidencias sin GTIN; no activar esta conducta en producción.
            product_ids_by_description = get_product_candidates_by_descriptions(
                line.description
                for invoice in invoices.values()
                for line in invoice.lines
                if line.description
            )
            print(f"Azure: {len(product_ids_by_barcode)} código(s) de barras asociado(s) a productos EFFI.")
            candidate_count = sum(len(products) for products in product_ids_by_description.values())
            print(f"Azure fallback por descripción: {candidate_count} candidato(s) evaluados.")
        except Exception as exc:
            print(f"ERROR al consultar Azure SQL: {exc}", file=sys.stderr)
            return 1

    for source in sources:
        if args.azure_lookup and source not in invoices:
            continue
        try:
            invoice = invoices[source] if args.azure_lookup else parse_invoice(source)
            destination = args.output_dir / (
                f"{_invoice_output_stem(source, invoice.invoice_id, invoice.supplier_tax_id)}_effi.xlsx"
            )
            warnings = generate_workbook(
                source,
                args.template,
                destination,
                tax_codes,
                args.json,
                product_ids_by_barcode,
                product_ids_by_description,
            )
            print(f"OK {source.name} -> {destination} ({len(invoice.lines)} líneas)")
            for warning in warnings:
                print(f"  AVISO: {warning}")
        except (InvoiceParseError, OSError, ValueError) as exc:
            failures += 1
            print(f"ERROR {source.name}: {exc}", file=sys.stderr)

    print(f"Proceso terminado: {len(sources) - failures} correcto(s), {failures} con error.")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())