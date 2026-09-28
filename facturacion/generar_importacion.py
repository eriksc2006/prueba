"""Parse DIAN UBL invoices and create EFFI purchase-detail import workbooks."""

from __future__ import annotations

import argparse
import json
import re
import sys
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable

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
    line_id: str
    item_code: str
    description: str
    quantity: Decimal
    unit: str
    unit_price: Decimal
    discount_total: Decimal
    net_amount: Decimal
    tax_rates: tuple[Decimal, ...]


@dataclass(frozen=True)
class Invoice:
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


def _line_warnings(line: InvoiceLine, mapping: dict[Decimal, int]) -> list[str]:
    warnings: list[str] = []
    label = f"línea {line.line_id or '?'}"
    if not line.item_code:
        warnings.append(f"{label}: no se encontró código de artículo; completar en EFFI.")
    if line.quantity <= 0:
        warnings.append(f"{label}: cantidad no positiva ({line.quantity}).")
    if len(line.tax_rates) != 1:
        warnings.append(f"{label}: tasa de impuesto ausente o múltiple; revisar código de impuesto.")
    elif _tax_code(line, mapping) is None:
        warnings.append(f"{label}: tasa {line.tax_rates[0]}% sin código EFFI configurado.")
    expected_net = line.unit_price * line.quantity - line.discount_total
    if abs(expected_net - line.net_amount) > PRICE_RECONCILIATION_LIMIT:
        warnings.append(
            f"{label}: precio × cantidad - descuento ({expected_net}) no coincide con el neto XML "
            f"({line.net_amount}); revisar la línea antes de importarla."
        )
    return warnings


def _write_detail(sheet: Worksheet, invoice: Invoice, tax_codes: dict[Decimal, int]) -> None:
    # The supplied workbook has sample lines in rows 2-4. Remove cell values
    # only; preserve the template's formatting and its header row.
    for row in sheet.iter_rows(min_row=2, max_row=max(sheet.max_row, len(invoice.lines) + 1), min_col=1, max_col=9):
        for cell in row:
            cell.value = None

    style_source = [sheet.cell(row=2, column=column) for column in range(1, 10)]
    for row_number, line in enumerate(invoice.lines, start=2):
        values = (
            line.item_code,             # A: EFFI article ID / barcode / serial
            None,                       # B: EFFI expense-type ID (not in DIAN UBL)
            None,                       # C: lot (not in DIAN UBL)
            None,                       # D: serial (not in DIAN UBL)
            None,                       # E: observation (kept for manual entry)
            line.quantity,              # F: quantity
            line.unit_price,            # G: unit price
            line.discount_total,        # H: total line discount
            _tax_code(line, tax_codes), # I: configured EFFI tax ID
        )
        for column, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_number, column=column, value=value)
            if row_number > 2:
                cell._style = style_source[column - 1]._style
                if style_source[column - 1].number_format:
                    cell.number_format = style_source[column - 1].number_format


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


def generate_workbook(
    source: Path,
    template: Path,
    output: Path,
    tax_codes: dict[Decimal, int],
    write_json: bool = False,
) -> list[str]:
    invoice = parse_invoice(source)
    workbook = load_workbook(template)
    sheet = workbook.active
    _write_detail(sheet, invoice, tax_codes)
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)
    workbook.close()

    if write_json:
        output.with_suffix(".json").write_text(
            json.dumps(_json_invoice(invoice), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return [warning for line in invoice.lines for warning in _line_warnings(line, tax_codes)]


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
    for source in sources:
        try:
            invoice = parse_invoice(source)
            destination = args.output_dir / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"
            warnings = generate_workbook(source, args.template, destination, tax_codes, args.json)
            print(f"OK {source.name} → {destination} ({len(invoice.lines)} líneas)")
            for warning in warnings:
                print(f"  AVISO: {warning}")
        except (InvoiceParseError, OSError, ValueError) as exc:
            failures += 1
            print(f"ERROR {source.name}: {exc}", file=sys.stderr)

    print(f"Proceso terminado: {len(sources) - failures} correcto(s), {failures} con error.")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())