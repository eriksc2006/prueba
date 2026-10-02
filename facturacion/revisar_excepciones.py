from __future__ import annotations

import html
import hashlib
import json
import os
import sys
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
from typing import Any

import streamlit as st
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
REPOSITORY_ROOT = ROOT.parent
if str(REPOSITORY_ROOT) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_ROOT))

from facturacion.azure_codes import search_products, search_products_by_barcode
from facturacion.generar_importacion import (
    DEFAULT_TAX_CODES,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE,
    DEFAULT_XML_DIR,
    _is_full_discount_bonus,
    _group_import_lines,
    _normalize_description,
    _presentation_tokens,
    _safe_stem,
    generate_workbook,
    parse_invoice,
)


REVIEW_FILE = DEFAULT_OUTPUT_DIR / "revisiones_manuales.json"
DETAIL_SHEET = "Hoja1"
INVOICE_APPROVAL_KEY = "__invoice_approval__"
INVOICE_APPROVAL_HISTORY_KEY = "__invoice_approval_history__"


def _display(value: Any, empty: str = "Sin dato") -> str:
    if value is None or str(value).strip().lower() in {"", "none"}:
        return empty
    return str(value)


def _money(value: Any) -> str:
    if value is None or str(value).strip().lower() in {"", "none"}:
        return "Sin dato"
    try:
        amount = Decimal(str(value))
        precision = max(2, -amount.as_tuple().exponent)
        formatted = f"{amount:,.{precision}f}"
        whole, _, fractional = formatted.partition(".")
        fractional = fractional.rstrip("0").ljust(2, "0")
        formatted = f"{whole}.{fractional}"
        localized = formatted.replace(",", "_").replace(".", ",").replace("_", ".")
        return f"COP $ {localized}"
    except Exception:
        return str(value)


def _load_review_state() -> dict[str, Any]:
    if not REVIEW_FILE.is_file():
        return {}
    try:
        return json.loads(REVIEW_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def _save_review_state(state: dict[str, Any]) -> None:
    REVIEW_FILE.parent.mkdir(parents=True, exist_ok=True)
    temporary = REVIEW_FILE.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    os.replace(temporary, REVIEW_FILE)


def _invoice_output_path(source: Path, invoice: Any) -> Path:
    return DEFAULT_OUTPUT_DIR / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"


@st.cache_data(ttl=900, show_spinner=False)
def _cached_barcode_product_ids(barcodes: tuple[str, ...]) -> dict[str, Any]:
    from facturacion.azure_codes import get_product_ids_by_barcodes

    return get_product_ids_by_barcodes(barcodes)


def _existing_tax_codes(
    output: Path,
    invoice: Any,
    product_ids_by_barcode: dict[str, Any],
    product_ids_by_description: dict[str, Any],
) -> dict[Decimal, int]:
    tax_codes = dict(DEFAULT_TAX_CODES)
    if not output.is_file():
        return tax_codes
    workbook = load_workbook(output, read_only=True, data_only=True)
    try:
        detail = workbook[DETAIL_SHEET] if DETAIL_SHEET in workbook.sheetnames else workbook.active
        observed: dict[Decimal, dict[int, int]] = {}
        groups = _group_import_lines(invoice, product_ids_by_barcode, product_ids_by_description)
        for row_number, lines in enumerate(groups, start=2):
            code_value = detail.cell(row=row_number, column=9).value
            if code_value is None:
                continue
            try:
                code = int(code_value)
            except (TypeError, ValueError):
                continue
            charged_rates = {
                rate
                for line in lines
                if line.net_amount != 0
                for rate in line.tax_rates
            }
            if len(charged_rates) == 1:
                rate = next(iter(charged_rates))
                rate_counts = observed.setdefault(rate, {})
                rate_counts[code] = rate_counts.get(code, 0) + 1
        for rate, code_counts in observed.items():
            tax_codes[rate] = max(code_counts, key=code_counts.get)
    finally:
        workbook.close()
    return tax_codes


def _sync_invoice_workbooks(force: bool = False) -> dict[str, Any]:
    parsed_invoices = []
    for source in sorted(DEFAULT_XML_DIR.glob("*.xml")):
        try:
            invoice = parse_invoice(source)
        except (OSError, ValueError):
            continue
        output = _invoice_output_path(source, invoice)
        try:
            source_is_newer = not output.is_file() or source.stat().st_mtime > output.stat().st_mtime
        except OSError:
            source_is_newer = True
        if force or source_is_newer:
            parsed_invoices.append((source, invoice, output))

    if not parsed_invoices:
        return {"generated": [], "mapped_barcodes": 0, "unmapped_barcodes": 0}

    invoices = [invoice for _, invoice, _ in parsed_invoices]
    barcode_codes = tuple(sorted({
        line.barcode for invoice in invoices for line in invoice.lines if line.barcode
    }))
    product_ids_by_barcode = _cached_barcode_product_ids(
        barcode_codes
    )
    product_ids_by_description: dict[str, Any] = {}
    generated_sources = []
    for source, invoice, output in parsed_invoices:
        tax_codes = _existing_tax_codes(
            output, invoice, product_ids_by_barcode, product_ids_by_description
        )
        generate_workbook(
            source,
            DEFAULT_TEMPLATE,
            output,
            tax_codes,
            product_ids_by_barcode=product_ids_by_barcode,
            product_ids_by_description=product_ids_by_description,
        )
        generated_sources.append(source.name)

    review_state = _load_review_state()
    approval_invalidated = False
    for source_name in generated_sources:
        invoice_reviews = review_state.get(source_name, {})
        if invoice_reviews.pop(INVOICE_APPROVAL_KEY, None) is not None:
            approval_invalidated = True
    if approval_invalidated:
        _save_review_state(review_state)
    _cached_invoice_queue.cache_clear()

    all_barcodes = [line.barcode for invoice in invoices for line in invoice.lines if line.barcode]
    mapped_barcodes = sum(barcode in product_ids_by_barcode for barcode in set(all_barcodes))
    return {
        "generated": generated_sources,
        "mapped_barcodes": mapped_barcodes,
        "unmapped_barcodes": len(set(all_barcodes)) - mapped_barcodes,
    }


def _invoice_review_complete(queue_item: dict[str, Any], invoice_reviews: dict[str, Any]) -> bool:
    return all(
        not _requires_human_decision(exception)
        or invoice_reviews.get(str(exception.get("Línea")), {}).get("status") in {"confirmed", "no_match"}
        for exception in queue_item["exceptions"]
    )


def _requires_human_decision(exception: dict[str, Any]) -> bool:
    reason = str(exception.get("Tipo") or "")
    if not reason:
        return True
    review_markers = (
        "no registrado en EFFI",
        "Sin coincidencia confiable",
        "Candidato cercano",
        "Neto XML conciliado",
        "Cantidad no positiva",
        "GTIN repetido con perfiles o unidades separadas",
    )
    if any(marker in reason for marker in review_markers):
        return True
    if "Bonificación 100% detectada" in reason and "Bonificación/prorrateo:" not in reason:
        return True
    return "Bonificación/prorrateo:" not in reason


def _approval_fingerprint(source: Path, invoice_reviews: dict[str, Any]) -> str:
    decisions = {
        line_id: {
            key: review.get(key)
            for key in ("status", "product_id", "final_unit_cost", "is_bonus", "bonus_link_line")
        }
        for line_id, review in sorted(invoice_reviews.items())
        if line_id not in {INVOICE_APPROVAL_KEY, INVOICE_APPROVAL_HISTORY_KEY} and isinstance(review, dict)
    }
    payload = {
        "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
        "decisions": decisions,
    }
    return hashlib.sha256(
        json.dumps(payload, ensure_ascii=False, sort_keys=True).encode("utf-8")
    ).hexdigest()


def _invoice_approval_is_current(source: Path, invoice_reviews: dict[str, Any]) -> bool:
    approval = invoice_reviews.get(INVOICE_APPROVAL_KEY, {})
    if approval.get("status") != "approved":
        return False
    try:
        return approval.get("fingerprint") == _approval_fingerprint(source, invoice_reviews)
    except OSError:
        return False


def _record_invoice_approval(
    state: dict[str, Any], source: Path, invoice_reviews: dict[str, Any], approver: str
) -> None:
    approval_history = invoice_reviews.setdefault(INVOICE_APPROVAL_HISTORY_KEY, [])
    previous_approval = invoice_reviews.get(INVOICE_APPROVAL_KEY)
    if previous_approval and (
        not approval_history or approval_history[-1].get("approved_at") != previous_approval.get("approved_at")
    ):
        approval_history.append(previous_approval)
    approval = {
        "status": "approved",
        "approved_by": approver.strip(),
        "approved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "fingerprint": _approval_fingerprint(source, invoice_reviews),
    }
    approval_history.append(approval)
    invoice_reviews[INVOICE_APPROVAL_KEY] = approval
    state[source.name] = invoice_reviews
    _save_review_state(state)


def _read_exception_rows(path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Excepciones" not in workbook.sheetnames:
            return []
        sheet = workbook["Excepciones"]
        rows = sheet.iter_rows(values_only=True)
        headers = [str(value).strip() if value is not None else "" for value in next(rows, ())]
        result = []
        for values in rows:
            item = dict(zip(headers, values))
            if item.get("Línea") is not None:
                result.append(item)
        return result
    finally:
        workbook.close()


def _invoice_queue() -> list[dict[str, Any]]:
    files = sorted(DEFAULT_XML_DIR.glob("*.xml")) + sorted(DEFAULT_OUTPUT_DIR.glob("*_effi.xlsx"))
    signature = []
    for path in files:
        try:
            metadata = path.stat()
            signature.append((str(path), metadata.st_mtime_ns, metadata.st_size))
        except OSError:
            signature.append((str(path), -1, -1))
    return _cached_invoice_queue(tuple(signature))


@lru_cache(maxsize=8)
def _cached_invoice_queue(signature: tuple[tuple[str, int, int], ...]) -> list[dict[str, Any]]:
    return _build_invoice_queue()


def _build_invoice_queue() -> list[dict[str, Any]]:
    queue = []
    for source in sorted(DEFAULT_XML_DIR.glob("*.xml")):
        try:
            invoice = parse_invoice(source)
        except (OSError, ValueError):
            continue
        output = DEFAULT_OUTPUT_DIR / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"
        exceptions = _read_exception_rows(output) if output.is_file() else []
        import_row_by_line = {
            str(line.line_id): row for row, line in enumerate(invoice.lines, start=2)
        }
        source_groups = _group_import_lines(invoice, None, None)
        if not output.is_file():
            import_row_by_line = {
                str(line.line_id): row
                for row, group in enumerate(source_groups, start=2)
                for line in group
            }
        article_ids: dict[str, Any] = {}
        product_data_by_line: dict[str, dict[str, Any]] = {}
        if output.is_file():
            workbook = load_workbook(output, read_only=True, data_only=True)
            try:
                detail = workbook[DETAIL_SHEET] if DETAIL_SHEET in workbook.sheetnames else workbook.active
                mapped_line_ids: set[str] = set()
                used_rows: set[int] = set()
                for item in exceptions:
                    line_id = str(item.get("Línea"))
                    try:
                        import_row = int(item.get("Fila importación"))
                    except (TypeError, ValueError):
                        import_row = None
                    if import_row and import_row >= 2:
                        import_row_by_line[line_id] = import_row
                        mapped_line_ids.add(line_id)
                        used_rows.add(import_row)

                remaining_rows = iter(
                    row for row in range(2, detail.max_row + 1) if row not in used_rows
                )
                for line in invoice.lines:
                    line_id = str(line.line_id)
                    if line_id not in mapped_line_ids:
                        import_row_by_line[line_id] = next(remaining_rows, import_row_by_line[line_id])
                product_data_by_line = {
                    line_id: {
                        "import_row": row,
                        "article_id": detail.cell(row=row, column=1).value,
                        "quantity": detail.cell(row=row, column=6).value,
                        "unit_price": detail.cell(row=row, column=7).value,
                        "discount": detail.cell(row=row, column=8).value,
                        "tax_code": detail.cell(row=row, column=9).value,
                    }
                    for line_id, row in import_row_by_line.items()
                }
                article_ids = {
                    line_id: values["article_id"]
                    for line_id, values in product_data_by_line.items()
                }
                for item in exceptions:
                    item["Artículo EFFI"] = article_ids.get(str(item.get("Línea")))
            finally:
                workbook.close()
        else:
            grouped_lines: dict[int, list[Any]] = {}
            for line in invoice.lines:
                grouped_lines.setdefault(import_row_by_line[str(line.line_id)], []).append(line)
            for import_row, lines in grouped_lines.items():
                quantity = sum((line.quantity for line in lines), Decimal("0"))
                net_total = sum((line.net_amount for line in lines), Decimal("0"))
                is_grouped = len(lines) > 1
                article_id = lines[0].item_code or None
                for line in lines:
                    product_data_by_line[str(line.line_id)] = {
                        "import_row": import_row,
                        "article_id": article_id,
                        "quantity": quantity,
                        "unit_price": net_total / quantity if is_grouped and quantity else line.unit_price,
                        "discount": Decimal("0") if is_grouped else line.discount_total,
                        "tax_code": None,
                    }
            article_ids = {
                line_id: values["article_id"]
                for line_id, values in product_data_by_line.items()
            }
        exceptions_by_line = {str(item.get("Línea")): item for item in exceptions}
        queue.append({
            "source": source,
            "invoice": invoice,
            "output": output if output.is_file() else None,
            "product_count": len(set(import_row_by_line.values())),
            "exceptions": exceptions,
            "exceptions_by_line": exceptions_by_line,
            "import_row_by_line": import_row_by_line,
            "article_ids": article_ids,
            "product_data_by_line": product_data_by_line,
        })
    return queue


def _product_from_exception(item: dict[str, Any]) -> dict[str, Any] | None:
    product_id = item.get("IdProducto candidato")
    if product_id is None:
        return None
    return {
        "IdProducto": product_id,
        "Nombre": item.get("Nombre candidato"),
        "Referencia": item.get("Referencia"),
        "Marca": item.get("Marca"),
        "Categoria": item.get("Categoría"),
        "PrecioConImpuestoTarifaNormal": item.get("Precio normal"),
        "PrecioConImpuestoTarifaMayorista": item.get("Precio mayorista"),
    }


def _product_label(product: dict[str, Any]) -> str:
    return (
        f"{product.get('IdProducto')} | {_display(product.get('Nombre'))} | "
        f"Ref. {_display(product.get('Referencia'))} | {_display(product.get('Marca'))}"
    )


def _format_presentation(value: Any) -> str:
    if value is None or not str(value).strip() or str(value).strip().lower() == "none":
        return "No identificada"
    return str(value)


def _record_review(
    state: dict[str, Any],
    source: Path,
    line_id: str,
    status: str,
    product: dict[str, Any] | None = None,
    reviewed_by: str = "",
) -> None:
    invoice_reviews = state.setdefault(source.name, {})
    invoice_reviews[line_id] = {
        "status": status,
        "product_id": product.get("IdProducto") if product else None,
        "product_name": product.get("Nombre") if product else None,
        "reference": product.get("Referencia") if product else None,
        "reviewed_by": reviewed_by.strip(),
        "reviewed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    _save_review_state(state)


def _record_cost_review(
    state: dict[str, Any],
    source: Path,
    line_id: str,
    final_unit_cost: Decimal,
    is_bonus: bool,
    bonus_link_line: str | None,
    reviewed_by: str = "",
) -> None:
    invoice_reviews = state.setdefault(source.name, {})
    line_review = invoice_reviews.setdefault(line_id, {})
    line_review.update({
        "final_unit_cost": str(final_unit_cost),
        "is_bonus": is_bonus,
        "bonus_link_line": bonus_link_line if is_bonus else None,
        "cost_saved_by": reviewed_by.strip(),
        "cost_saved_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    })
    if line_review.get("status") not in {"confirmed", "no_match"}:
        line_review["status"] = "cost_saved"
    _save_review_state(state)


def _line_product_key(line: Any) -> tuple[str, str]:
    if line.barcode:
        return "barcode", line.barcode
    if line.item_code and line.item_code not in {"0", "-"}:
        return "item", line.item_code.strip().casefold()
    return "description", _normalize_description(line.description)


def _suggest_bonus_link(line: Any, invoice: Any) -> str | None:
    if line.net_amount != 0:
        return None
    key = _line_product_key(line)
    return next(
        (
            other.line_id
            for other in invoice.lines
            if other.line_id != line.line_id
            and _line_product_key(other) == key
            and other.net_amount > 0
        ),
        None,
    )


def _bonus_cost_summary(source_lines: list[Any], bonus_line_ids: set[str]) -> dict[str, Decimal]:
    unique_lines = list({str(line.line_id): line for line in source_lines}.values())
    if any(
        str(line.line_id) in bonus_line_ids and line.net_amount != 0
        for line in unique_lines
    ):
        raise ValueError("La línea marcada como bonificación debe tener neto cero en la factura.")
    total_quantity = sum((line.quantity for line in unique_lines), Decimal("0"))
    net_total = sum((line.net_amount for line in unique_lines), Decimal("0"))
    billed_quantity = sum(
        (
            line.quantity
            for line in unique_lines
            if str(line.line_id) not in bonus_line_ids and line.net_amount > 0
        ),
        Decimal("0"),
    )
    return {
        "billed_quantity": billed_quantity,
        "bonus_quantity": total_quantity - billed_quantity,
        "total_quantity": total_quantity,
        "net_total": net_total,
        "effective_unit_cost": net_total / total_quantity if total_quantity > 0 else Decimal("0"),
    }


def _render_cost_editor(
    source: Path,
    invoice: Any,
    line: Any,
    invoice_reviews: dict[str, Any],
    review_state: dict[str, Any],
    line_by_id: dict[str, Any],
    product_row_by_line: dict[str, dict[str, Any]],
    reviewer_name: str,
) -> None:
    line_id = str(line.line_id)
    saved_review = invoice_reviews.get(line_id, {})
    product_row = product_row_by_line[line_id]
    default_cost = Decimal(str(saved_review.get("final_unit_cost", product_row["effective_cost"])))
    suggested_link = _suggest_bonus_link(line, invoice)
    default_bonus = saved_review.get("is_bonus", line.net_amount == 0 and suggested_link is not None)
    paid_line_ids = [
        str(paid_line.line_id)
        for paid_line in invoice.lines
        if str(paid_line.line_id) != line_id
        and paid_line.net_amount > 0
        and not invoice_reviews.get(str(paid_line.line_id), {}).get("is_bonus")
    ]
    saved_link = saved_review.get("bonus_link_line") or suggested_link
    default_link_index = paid_line_ids.index(saved_link) if saved_link in paid_line_ids else 0

    st.markdown("#### Costo y bonificación")
    is_bonus = st.checkbox(
        "Es bonificación",
        value=bool(default_bonus),
        key=f"is_bonus_{source.name}_{line_id}",
        help="Marca las unidades gratuitas para sumarlas al producto pagado que las financia.",
    )
    bonus_link_line = None
    allocation = None
    validation_error = ""
    if is_bonus and paid_line_ids:
        bonus_link_line = st.selectbox(
            "Bonificada sobre",
            paid_line_ids,
            index=default_link_index,
            format_func=lambda paid_id: (
                f"Línea {paid_id} · {line_by_id[paid_id].description or line_by_id[paid_id].item_code}"
            ),
            key=f"bonus_link_{source.name}_{line_id}",
        )
        paid_row = product_row_by_line[bonus_link_line]
        bonus_line_ids = {
            str(bonus_line.line_id)
            for bonus_line in invoice.lines
            if invoice_reviews.get(str(bonus_line.line_id), {}).get("is_bonus")
            and invoice_reviews.get(str(bonus_line.line_id), {}).get("bonus_link_line") == bonus_link_line
        }
        bonus_line_ids.add(line_id)
        source_ids = list(dict.fromkeys(paid_row["line_ids"] + list(bonus_line_ids)))
        source_lines = [line_by_id[source_id] for source_id in source_ids if source_id in line_by_id]
        try:
            allocation = _bonus_cost_summary(source_lines, bonus_line_ids)
        except ValueError as error:
            validation_error = str(error)
        if allocation:
            profiles = {
                tuple(rate.normalize() for rate in source_line.tax_rates)
                for source_line in source_lines
                if source_line.net_amount != 0
            }
            if len(profiles) > 1:
                validation_error = "El producto pagado y sus bonificaciones deben compartir impuesto."
            elif allocation["total_quantity"] <= 0 or allocation["net_total"] < 0:
                validation_error = "Las cantidades o el neto del grupo no son válidos."
    elif is_bonus:
        validation_error = "No hay producto pagado disponible para vincular la bonificación."

    output_cost = allocation["effective_unit_cost"] if is_bonus and allocation else default_cost
    cost_column, summary_column = st.columns([1.15, 1])
    with cost_column:
        final_cost = st.number_input(
            "Costo unitario (COP)" if not is_bonus else "Costo prorrateado (COP / und.)",
            min_value=0.0,
            value=float(output_cost),
            step=0.01,
            format="%.6f",
            key=(f"bonus_cost_{source.name}_{line_id}_{bonus_link_line}" if is_bonus else f"regular_cost_{source.name}_{line_id}"),
            disabled=is_bonus,
        )
    with summary_column:
        if is_bonus and allocation:
            st.metric("Unidades del grupo", allocation["total_quantity"])
            st.caption(f'{allocation["billed_quantity"]} pagadas + {allocation["bonus_quantity"]} gratis')
        else:
            st.metric("Neto de línea", _money(product_row["net_total"]))

    if is_bonus and allocation:
        calculated_net = (allocation["effective_unit_cost"] * allocation["total_quantity"]).quantize(Decimal("0.01"))
        st.caption(
            f'Neto del grupo {_money(allocation["net_total"])} ÷ {allocation["total_quantity"]} '
            f'und. = {_money(allocation["effective_unit_cost"])} / und.'
        )
        if abs(calculated_net - allocation["net_total"].quantize(Decimal("0.01"))) <= Decimal("0.02"):
            st.success("El neto del grupo se conserva.")
    elif not is_bonus:
        calculated_net = (Decimal(str(final_cost)) * product_row["quantity"]).quantize(Decimal("0.01"))
        if abs(calculated_net - product_row["net_total"].quantize(Decimal("0.01"))) <= Decimal("0.02"):
            st.caption(f"Neto calculado {_money(calculated_net)} · coincide con factura")
        else:
            st.warning(f"Neto calculado {_money(calculated_net)} · factura {_money(product_row['net_total'])}")
    if validation_error:
        st.error(validation_error)
    if st.button(
        "Guardar costo",
        type="primary",
        key=f"save_cost_{source.name}_{line_id}",
        disabled=is_bonus and bool(validation_error),
    ):
        _record_cost_review(
            review_state,
            source,
            line_id,
            allocation["effective_unit_cost"] if is_bonus and allocation else Decimal(str(final_cost)),
            is_bonus,
            bonus_link_line,
            reviewer_name,
        )
        st.success("Costo y bonificación guardados.")
        st.rerun()


def _invoice_product_rows(
    queue_item: dict[str, Any], invoice_reviews: dict[str, Any] | None = None
) -> list[dict[str, Any]]:
    invoice = queue_item["invoice"]
    invoice_reviews = invoice_reviews or {}
    rows_by_import_row: dict[int, list[Any]] = {}
    for line in invoice.lines:
        row = queue_item["import_row_by_line"].get(str(line.line_id))
        if row is not None:
            rows_by_import_row.setdefault(row, []).append(line)

    result = []
    for import_row, lines in sorted(rows_by_import_row.items()):
        line_ids = [str(line.line_id) for line in lines]
        total_quantity = sum((line.quantity for line in lines), Decimal("0"))
        net_total = sum((line.net_amount for line in lines), Decimal("0"))
        effective_cost = net_total / total_quantity if total_quantity > 0 else Decimal("0")
        representative_id = line_ids[0]
        data = queue_item["product_data_by_line"].get(representative_id, {})
        reviewed_costs = [
            Decimal(str(invoice_reviews[line_id]["final_unit_cost"]))
            for line_id in line_ids
            if line_id in invoice_reviews and invoice_reviews[line_id].get("final_unit_cost") is not None
        ]
        if reviewed_costs:
            effective_cost = reviewed_costs[-1]
        result.append({
            "import_row": import_row,
            "line_ids": line_ids,
            "lines": lines,
            "description": " / ".join(dict.fromkeys(line.description for line in lines if line.description)),
            "barcode": next((line.barcode for line in lines if line.barcode), ""),
            "item_code": next((line.item_code for line in lines if line.item_code), ""),
            "article_id": data.get("article_id"),
            "quantity": total_quantity,
            "net_total": net_total,
            "effective_cost": effective_cost,
            "tax_code": data.get("tax_code"),
            "bonus": (
                any(invoice_reviews.get(line_id, {}).get("is_bonus") for line_id in line_ids)
                or (any(line.net_amount == 0 for line in lines) and len(lines) > 1)
                or any(_is_full_discount_bonus(line) for line in lines)
            ),
            "bonus_link_line": next((
                invoice_reviews[line_id].get("bonus_link_line")
                for line_id in line_ids
                if invoice_reviews.get(line_id, {}).get("bonus_link_line")
            ), None),
        })
    return result


def _product_review_status(
    product_row: dict[str, Any],
    exceptions_by_line: dict[str, dict[str, Any]],
    invoice_reviews: dict[str, Any],
) -> str:
    reviewable_ids = [
        line_id
        for line_id in product_row["line_ids"]
        if line_id in exceptions_by_line
        and _requires_human_decision(exceptions_by_line[line_id])
    ]
    if not reviewable_ids:
        return "Validado automáticamente"
    completed = all(
        invoice_reviews.get(line_id, {}).get("status") in {"confirmed", "no_match"}
        for line_id in reviewable_ids
    )
    return "Revisado por persona" if completed else "Requiere revisión"


def _pending_review_count(queue_item: dict[str, Any], invoice_reviews: dict[str, Any]) -> int:
    exceptions_by_line = queue_item["exceptions_by_line"]
    return sum(
        _product_review_status(product_row, exceptions_by_line, invoice_reviews) == "Requiere revisión"
        for product_row in _invoice_product_rows(queue_item, invoice_reviews)
    )


def _review_status(queue_item: dict[str, Any], invoice_reviews: dict[str, Any]) -> str:
    if _invoice_approval_is_current(queue_item["source"], invoice_reviews):
        return "Revisada"
    pending = _pending_review_count(queue_item, invoice_reviews)
    if not pending:
        return "Pendiente de aprobación"
    has_decision = any(
        invoice_reviews.get(str(exception.get("Línea")), {}).get("status") in {"confirmed", "no_match"}
        for exception in queue_item["exceptions"]
    )
    return "En revisión" if has_decision else "Pendiente"


def _dashboard_metrics(queue: list[dict[str, Any]], review_state: dict[str, Any]) -> dict[str, int]:
    pending_invoice_count = 0
    in_progress_invoice_count = 0
    pending_product_count = 0
    gtin_exception_count = 0
    reviewed_today_count = 0
    today = datetime.now().astimezone().date().isoformat()
    for item in queue:
        invoice_reviews = review_state.get(item["source"].name, {})
        pending = _pending_review_count(item, invoice_reviews)
        pending_product_count += pending
        pending_invoice_count += pending > 0
        in_progress_invoice_count += _review_status(item, invoice_reviews) == "En revisión"
        gtin_exception_count += sum(
            bool(line.barcode) and not item["article_ids"].get(str(line.line_id))
            for line in item["invoice"].lines
        )
        invoice_reviewed_today = any(
            str(review.get(timestamp_key, "")).startswith(today)
            for line_id, review in invoice_reviews.items()
            if line_id != INVOICE_APPROVAL_KEY and isinstance(review, dict)
            for timestamp_key in ("reviewed_at", "cost_saved_at")
        )
        approval = invoice_reviews.get(INVOICE_APPROVAL_KEY, {})
        invoice_reviewed_today |= str(approval.get("approved_at", "")).startswith(today)
        reviewed_today_count += invoice_reviewed_today
    return {
        "pending_invoices": pending_invoice_count,
        "in_progress_invoices": in_progress_invoice_count,
        "pending_products": pending_product_count,
        "gtin_exceptions": gtin_exception_count,
        "reviewed_today": reviewed_today_count,
    }


def _product_review_rows(
    queue: list[dict[str, Any]], review_state: dict[str, Any]
) -> list[dict[str, Any]]:
    result = []
    for invoice_index, item in enumerate(queue):
        source = item["source"]
        invoice = item["invoice"]
        invoice_reviews = review_state.get(source.name, {})
        line_by_id = {str(line.line_id): line for line in invoice.lines}
        for exception_index, exception in enumerate(item["exceptions"]):
            line_id = str(exception.get("Línea"))
            line = line_by_id.get(line_id)
            review = invoice_reviews.get(line_id, {})
            status = review.get("status")
            result.append({
                "Factura": invoice.invoice_id,
                "Proveedor": invoice.supplier_name or source.name,
                "NIT": invoice.supplier_tax_id,
                "Línea": line_id,
                "Producto": line.description if line else exception.get("Descripción factura"),
                "Código de barras": line.barcode if line else exception.get("Código de barras"),
                "Estado": (
                    "Confirmado" if status == "confirmed"
                    else "Sin coincidencia" if status == "no_match"
                    else "Automático" if not _requires_human_decision(exception)
                    else "Pendiente"
                ),
                "Decisión": review.get("product_name") or ("Sin coincidencia" if status == "no_match" else "Por decidir"),
                "_invoice_index": invoice_index,
                "_exception_index": exception_index,
                "_source_name": source.name,
            })
    return result


def _history_rows(queue: list[dict[str, Any]], review_state: dict[str, Any]) -> list[dict[str, Any]]:
    queue_by_source = {item["source"].name: item for item in queue}
    rows = []
    for source_name, invoice_reviews in review_state.items():
        item = queue_by_source.get(source_name)
        invoice_id = item["invoice"].invoice_id if item else source_name
        supplier = item["invoice"].supplier_name if item else "Sin proveedor"
        for line_id, review in invoice_reviews.items():
            if line_id == INVOICE_APPROVAL_KEY or not isinstance(review, dict):
                continue
            status = review.get("status")
            timestamp = review.get("reviewed_at") or review.get("cost_saved_at")
            if status not in {"confirmed", "no_match"} and not timestamp:
                continue
            rows.append({
                "Fecha": timestamp or "",
                "Factura": invoice_id,
                "Proveedor": supplier or "Sin proveedor",
                "Línea": line_id,
                "Decisión": review.get("product_name") or ("Sin coincidencia" if status == "no_match" else "Costo / bonificación"),
                "Estado": "Confirmado" if status == "confirmed" else "Sin coincidencia" if status == "no_match" else "Costo guardado",
                "Responsable": review.get("reviewed_by") or review.get("cost_saved_by", ""),
            })
        approval = invoice_reviews.get(INVOICE_APPROVAL_KEY, {})
        approval_history = invoice_reviews.get(INVOICE_APPROVAL_HISTORY_KEY, [])
        if not approval_history and approval.get("status") == "approved":
            approval_history = [approval]
        approval_is_current = bool(item and _invoice_approval_is_current(item["source"], invoice_reviews))
        for approval_record in approval_history:
            if approval_record.get("status") != "approved":
                continue
            is_current = (
                approval_is_current
                and approval_record.get("approved_at") == approval.get("approved_at")
            )
            rows.append({
                "Fecha": approval_record.get("approved_at", ""),
                "Factura": invoice_id,
                "Proveedor": supplier or "Sin proveedor",
                "Línea": "—",
                "Decisión": "Importación aprobada",
                "Estado": "Aprobada" if is_current else "Aprobación anterior",
                "Responsable": approval_record.get("approved_by", ""),
            })
    return sorted(rows, key=lambda row: row["Fecha"], reverse=True)


def _open_product_review(invoice_index: int, exception_index: int, source_name: str) -> None:
    st.session_state["main_view"] = "Facturas"
    st.session_state["review_invoice_selection"] = invoice_index
    st.session_state[f"review_cursor_{source_name}"] = exception_index


def _select_invoice(invoice_index: int) -> None:
    st.session_state["review_invoice_selection"] = invoice_index


def _next_pending_index(
    exceptions: list[dict[str, Any]],
    invoice_reviews: dict[str, Any],
    start_index: int,
) -> int | None:
    if not exceptions:
        return None
    indexes = list(range(start_index, len(exceptions))) + list(range(0, min(start_index, len(exceptions))))
    for index in indexes:
        exception = exceptions[index]
        if not _requires_human_decision(exception):
            continue
        line_id = str(exception.get("Línea"))
        if invoice_reviews.get(line_id, {}).get("status") not in {"confirmed", "no_match"}:
            return index
    return None


def _next_pending_invoice_index(
    queue: list[dict[str, Any]],
    review_state: dict[str, Any],
    start_index: int,
) -> int | None:
    if not queue:
        return None
    indexes = list(range(start_index, len(queue))) + list(range(0, min(start_index, len(queue))))
    for index in indexes:
        item = queue[index]
        source_name = item["source"].name
        if _next_pending_index(item["exceptions"], review_state.get(source_name, {}), 0) is not None:
            return index
    return None


def _reviewed_workbook(queue_item: dict[str, Any], approved: dict[str, Any]) -> Path:
    invoice = queue_item["invoice"]
    source = queue_item["source"]
    original = queue_item["output"]
    if original is None or not original.is_file():
        from facturacion.generar_importacion import DEFAULT_TAX_CODES, DEFAULT_TEMPLATE, generate_workbook

        original = DEFAULT_OUTPUT_DIR / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"
        generate_workbook(source, DEFAULT_TEMPLATE, original, DEFAULT_TAX_CODES)
    workbook = load_workbook(original)
    sheet = workbook[DETAIL_SHEET] if DETAIL_SHEET in workbook.sheetnames else workbook.active
    line_rows = {
        str(line_id): int(row)
        for line_id, row in queue_item.get("import_row_by_line", {}).items()
    }
    if not line_rows:
        line_rows = {str(line.line_id): row for row, line in enumerate(invoice.lines, start=2)}
    for exception in _read_exception_rows(original):
        try:
            import_row = int(exception.get("Fila importación"))
        except (TypeError, ValueError):
            continue
        if import_row >= 2:
            line_rows[str(exception.get("Línea"))] = import_row
    row_reviews: dict[int, dict[str, Any]] = {}
    for line_id, review in approved.items():
        row = line_rows.get(str(line_id))
        if row is None:
            continue
        if review.get("status") == "confirmed" and review.get("product_id") is not None:
            existing = row_reviews.get(row)
            if existing and existing.get("status") == "confirmed":
                if existing.get("product_id") != review["product_id"]:
                    raise ValueError("Las líneas agrupadas de un producto tienen artículos EFFI diferentes.")
            row_reviews[row] = review
        elif review.get("status") == "no_match":
            if row not in row_reviews or row_reviews[row].get("status") != "confirmed":
                row_reviews[row] = review
    for row, review in row_reviews.items():
        if review.get("status") == "confirmed":
            sheet.cell(row=row, column=1, value=review["product_id"])
        elif review.get("status") == "no_match":
            sheet.cell(row=row, column=1).value = None

    line_by_id = {str(line.line_id): line for line in invoice.lines}
    line_ids_by_row: dict[int, list[str]] = {}
    for line_id, row in line_rows.items():
        line_ids_by_row.setdefault(row, []).append(line_id)

    bonus_links: dict[int, list[str]] = {}
    bonus_rows: set[int] = set()
    for line_id, review in approved.items():
        if not review.get("is_bonus"):
            continue
        target_line_id = str(review.get("bonus_link_line") or "")
        target_row = line_rows.get(target_line_id)
        bonus_row = line_rows.get(str(line_id))
        if target_row is None or bonus_row is None or target_line_id == str(line_id):
            raise ValueError(f"La bonificación de la línea {line_id} no tiene un producto pagado válido vinculado.")
        bonus_links.setdefault(target_row, []).append(str(line_id))
        if bonus_row != target_row:
            bonus_rows.add(bonus_row)

    def reviewed_cost(line_ids: list[str]) -> Decimal | None:
        overrides = [
            (str(approved[line_id].get("cost_saved_at") or ""), Decimal(str(approved[line_id]["final_unit_cost"])))
            for line_id in line_ids
            if line_id in approved and approved[line_id].get("final_unit_cost") is not None
        ]
        return max(overrides, key=lambda value: value[0])[1] if overrides else None

    bonus_target_rows = set(bonus_links)
    for target_row, bonus_line_ids in bonus_links.items():
        source_line_ids = list(dict.fromkeys(
            line_ids_by_row.get(target_row, []) + bonus_line_ids
        ))
        source_lines = [line_by_id[line_id] for line_id in source_line_ids if line_id in line_by_id]
        charged_tax_profiles = {
            tuple(rate.normalize() for rate in line.tax_rates)
            for line in source_lines
            if line.net_amount != 0
        }
        if len(charged_tax_profiles) > 1:
            raise ValueError("No se pueden agrupar bonificaciones con perfiles de impuesto pagados diferentes.")
        allocation = _bonus_cost_summary(source_lines, set(bonus_line_ids))
        total_quantity = allocation["total_quantity"]
        net_total = allocation["net_total"]
        if total_quantity <= 0 or net_total < 0:
            raise ValueError("Las cantidades o el neto del grupo de bonificación no son válidos.")
        cost_override = reviewed_cost(source_line_ids)
        unit_cost = cost_override if cost_override is not None else allocation["effective_unit_cost"]
        sheet.cell(row=target_row, column=6, value=total_quantity)
        sheet.cell(row=target_row, column=7, value=unit_cost)
        sheet.cell(row=target_row, column=8, value=Decimal("0"))
        for bonus_line_id in bonus_line_ids:
            bonus_row = line_rows[bonus_line_id]
            if bonus_row != target_row:
                for column in range(1, 10):
                    sheet.cell(row=bonus_row, column=column).value = None

    for row, line_ids in line_ids_by_row.items():
        if row in bonus_target_rows or row in bonus_rows:
            continue
        unit_cost = reviewed_cost(line_ids)
        if unit_cost is None:
            continue
        source_lines = [line_by_id[line_id] for line_id in line_ids if line_id in line_by_id]
        total_quantity = sum((line.quantity for line in source_lines), Decimal("0"))
        sheet.cell(row=row, column=6, value=total_quantity)
        sheet.cell(row=row, column=7, value=unit_cost)
        sheet.cell(row=row, column=8, value=Decimal("0"))
    destination = original.with_name(original.stem.removesuffix("_effi") + "_revisado.xlsx")
    workbook.save(destination)
    workbook.close()
    return destination


def _render_product(product: dict[str, Any] | None, confidence: Any = None) -> None:
    if not product:
        st.info("No hay producto sugerido. Usa la búsqueda manual en EFFI.")
        return
    confidence_text = f"{float(confidence):.0%}" if confidence is not None else "Selección manual"
    st.markdown(f"**{_display(product.get('Nombre'))}**")
    st.caption(
        f"EFFI {product.get('IdProducto')} · Ref. {_display(product.get('Referencia'))} · "
        f"Marca {_display(product.get('Marca'))} · Categoría {_display(product.get('Categoria'))}"
    )
    product_presentation = _presentation_tokens(
        str(product.get("Nombre") or "") + " " + str(product.get("Referencia") or "")
    )
    st.caption(
        "Presentación EFFI: " + (", ".join(sorted(product_presentation)) if product_presentation else "No informada")
    )
    left, right = st.columns(2)
    left.metric("Precio normal", _money(product.get("PrecioConImpuestoTarifaNormal")))
    right.metric("Precio mayorista", _money(product.get("PrecioConImpuestoTarifaMayorista")))
    st.caption(f"Similitud del texto: {confidence_text}")


def _render_styles() -> None:
    st.markdown("""
    <style>
            :root { --ink: #202b27; --paper: #f4f6f4; --line: #dce3df; --green: #1c5b4b; --muted: #697570; }
            .stApp { background: var(--paper); color: var(--ink); }
            [data-testid="stHeader"] { background: #fff; border-bottom: 1px solid var(--line); }
            [data-testid="stSidebar"] { background: #183b32; }
            [data-testid="stSidebar"] * { color: #f4f7f5; }
            [data-testid="stSidebar"] input { background: #214b40; border-color: rgba(244,247,245,.25); }
            [data-testid="stSidebar"] [data-testid="stRadio"] label { align-items: flex-start; border-bottom: 1px solid rgba(244,247,245,.16); padding: .62rem .25rem; }
            [data-testid="stSidebar"] [data-testid="stRadio"] label p { color: #f4f7f5 !important; line-height: 1.3; overflow-wrap: anywhere; white-space: normal !important; }
            [data-testid="stMetricValue"] { font-size: 1.08rem; overflow-wrap: anywhere; white-space: normal; }
            .block-container { max-width: 1240px; padding-top: .2rem; padding-bottom: 2rem; }
            .app-topbar { height: 3rem; display: grid; grid-template-columns: 1fr 1fr 1fr; align-items: center; border-bottom: 1px solid var(--line); margin: 0 -1rem 1.15rem; padding: 0 1rem; background: #fff; }
            .app-topbar-title { text-align: center; font-weight: 600; font-size: .92rem; }
            .app-topbar-user { justify-self: end; width: 2rem; height: 2rem; display: grid; place-items: center; border: 1px solid var(--line); border-radius: 50%; color: var(--green); font-weight: 700; }
            .sidebar-brand, .sidebar-profile { display: flex; align-items: center; gap: .7rem; padding: .6rem .2rem 1rem; }
            .brand-mark, .profile-mark { width: 2.25rem; height: 2.25rem; display: grid; place-items: center; border-radius: 6px; background: #dcebe4; color: #183b32 !important; font-weight: 700; }
            .sidebar-brand strong, .sidebar-brand small, .sidebar-profile strong, .sidebar-profile small { display: block; }
            .sidebar-brand small, .sidebar-profile small { color: #b9cec5; font-size: .72rem; margin-top: .12rem; }
            .sidebar-status { display: grid; gap: .24rem; margin: 1.1rem 0 .75rem; padding: .8rem; border: 1px solid rgba(244,247,245,.22); border-radius: 5px; }
            .sidebar-status small, .sidebar-status span { color: #b9cec5; font-size: .72rem; }
            .sidebar-status strong { color: #e8f2ed; font-size: .84rem; }
            .sidebar-profile { margin-top: .8rem; border-top: 1px solid rgba(244,247,245,.2); padding-top: 1rem; }
            .page-eyebrow { color: #437e69; font-size: .69rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; margin: .45rem 0 .25rem; }
            h1, h2, h3 { color: var(--ink); font-family: "DM Sans", "Segoe UI", sans-serif; letter-spacing: 0; }
            h1 { font-size: 1.65rem; line-height: 1.2; margin: 0 0 .3rem; }
            h2 { font-size: 1.12rem; }
            h3, h4 { font-size: 1rem; }
            .page-subtitle { color: var(--muted); font-size: .92rem; margin: 0 0 1rem; }
            .summary-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); background: #fff; border: 1px solid var(--line); border-radius: 6px; margin: .5rem 0 1.4rem; }
            .summary-card { min-height: 6.3rem; padding: 1rem .85rem .8rem; }
            .summary-card + .summary-card { border-left: 1px solid var(--line); }
            .summary-label { color: var(--muted); font-size: .78rem; margin-bottom: .25rem; }
            .summary-value { color: var(--ink); font-size: 1.75rem; font-weight: 650; line-height: 1.2; }
            .summary-note { color: #7b8580; font-size: .72rem; margin-top: .55rem; }
            .section-title { font-size: 1.18rem; font-weight: 650; margin: 1.25rem 0 .25rem; }
            .compare-panel { border-top: 3px solid var(--green); padding: .45rem 0 .15rem; min-height: 0; }
            .compare-panel.xml { border-color: #c45135; }
            div.stButton > button { min-height: 2.55rem; border-radius: 5px; border-color: var(--line); }
            div.stButton > button[kind="primary"] { background: var(--green); border-color: var(--green); color: #fff; }
            [data-testid="stMetric"] { background: transparent; border-bottom: 1px solid var(--line); padding: .2rem 0; }
            div[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 5px; }
            @media (max-width: 900px) {
                .summary-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
                .summary-card:nth-child(3) { border-left: 0; border-top: 1px solid var(--line); }
                .summary-card:nth-child(4) { border-top: 1px solid var(--line); }
            }
            @media (max-width: 640px) {
                .block-container { padding-left: .8rem; padding-right: .8rem; }
                .app-topbar { margin-left: -.8rem; margin-right: -.8rem; padding: 0 .8rem; }
                h1 { font-size: 1.55rem; }
                [data-testid="stHorizontalBlock"] { gap: .5rem; }
            }
    </style>
    """, unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(
        page_title="Revisión EFFI · Human on the loop",
        page_icon="E",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _render_styles()
    st.markdown(
        '<div class="app-topbar"><span></span><span class="app-topbar-title">Facturas</span>'
        '<span class="app-topbar-user" title="Revisión EFFI">E</span></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="page-eyebrow">Operación diaria</div>', unsafe_allow_html=True)
    st.markdown('<h1>Bandeja de facturas</h1>', unsafe_allow_html=True)
    st.markdown(
        '<p class="page-subtitle">Modo Human on the loop: el sistema propone y el revisor decide; la sincronización es manual para no bloquear la cola.</p>',
        unsafe_allow_html=True,
    )

    sync_notice = st.session_state.pop("invoice_sync_notice", None)
    sync_error = st.session_state.pop("invoice_sync_error", None)
    if sync_notice:
        st.success(sync_notice)
    if sync_error:
        st.warning(f"No se pudo consultar Azure para generar las facturas automáticamente: {sync_error}")

    queue = _invoice_queue()
    if not queue:
        st.info("No hay facturas descargadas. Coloca los XML de EFFI en la carpeta de entrada y genera los libros.")
        return

    review_state = _load_review_state()
    metrics = _dashboard_metrics(queue, review_state)
    section_counts = {
        "Facturas": metrics["pending_invoices"],
        "Revisión de productos": metrics["pending_products"],
        "Excepciones de GTIN": metrics["gtin_exceptions"],
    }
    st.sidebar.markdown(
        '<div class="sidebar-brand"><span class="brand-mark">E</span><span><strong>EFFI</strong>'
        '<small>Control de facturas</small></span></div>',
        unsafe_allow_html=True,
    )
    main_view = st.sidebar.radio(
        "Sección",
        ("Facturas", "Revisión de productos", "Excepciones de GTIN", "Historial"),
        key="main_view",
        format_func=lambda section: (
            f"{section}  {section_counts[section]}" if section_counts.get(section) else section
        ),
    )
    st.sidebar.markdown(
        '<div class="sidebar-status"><small>Sincronización EFFI</small>'
        '<strong>Datos XML locales</strong><span>Actualiza para volver a leer la bandeja</span></div>',
        unsafe_allow_html=True,
    )
    reviewer_name = st.sidebar.text_input(
        "Revisor activo",
        key="active_reviewer",
        placeholder="Tu nombre para aprobar",
    )
    st.sidebar.markdown(
        f'<div class="sidebar-profile"><span class="profile-mark">'
        f'{reviewer_name[:1].upper() if reviewer_name else "R"}</span><span><strong>'
        f'{reviewer_name or "Revisor"}</strong><small>Analista de facturación</small></span></div>',
        unsafe_allow_html=True,
    )

    if main_view == "Revisión de productos":
        st.markdown('<div class="page-eyebrow">Cola de trabajo</div>', unsafe_allow_html=True)
        st.markdown('<h1>Revisión de productos</h1>', unsafe_allow_html=True)
        st.markdown('<p class="page-subtitle">Resuelve las coincidencias y excepciones de cada línea.</p>', unsafe_allow_html=True)
        product_rows = _product_review_rows(queue, review_state)
        search_column, status_column, supplier_column = st.columns([2, 1, 1])
        search_text = search_column.text_input("Buscar producto, factura o código", key="product_queue_search")
        status_filter = status_column.selectbox("Estado", ("Todos", "Pendiente", "Automático", "Confirmado", "Sin coincidencia"), key="product_queue_status")
        supplier_options = ["Todos"] + sorted({row["Proveedor"] for row in product_rows})
        supplier_filter = supplier_column.selectbox("Proveedor", supplier_options, key="product_queue_supplier")
        normalized_search = search_text.strip().casefold()
        filtered_products = [
            row for row in product_rows
            if (status_filter == "Todos" or row["Estado"] == status_filter)
            and (supplier_filter == "Todos" or row["Proveedor"] == supplier_filter)
            and (not normalized_search or normalized_search in " ".join(
                str(row.get(key, "")) for key in ("Factura", "Proveedor", "NIT", "Producto", "Código de barras")
            ).casefold())
        ]
        st.dataframe(
            [{key: value for key, value in row.items() if not key.startswith("_")}
             for row in filtered_products],
            use_container_width=True,
            hide_index=True,
        )
        if filtered_products:
            product_index = st.selectbox(
                "Línea para revisar",
                range(len(filtered_products)),
                format_func=lambda index: (
                    f'{filtered_products[index]["Factura"]} · Línea {filtered_products[index]["Línea"]} · '
                    f'{filtered_products[index]["Producto"]}'
                ),
                key="product_queue_selection",
            )
            selected_product = filtered_products[product_index]
            st.button(
                "Abrir revisión",
                type="primary",
                on_click=_open_product_review,
                args=(
                    selected_product["_invoice_index"],
                    selected_product["_exception_index"],
                    selected_product["_source_name"],
                ),
            )
        else:
            st.info("No hay productos que coincidan con estos filtros.")
        return

    if main_view == "Excepciones de GTIN":
        st.markdown('<div class="page-eyebrow">Acción en catálogo EFFI</div>', unsafe_allow_html=True)
        st.markdown('<h1>Excepciones de GTIN</h1>', unsafe_allow_html=True)
        st.markdown('<p class="page-subtitle">Códigos de barras de factura que aún no están vinculados a un artículo.</p>', unsafe_allow_html=True)
        barcode_rows = []
        for invoice_index, queued_invoice in enumerate(queue):
            for line in queued_invoice["invoice"].lines:
                line_id = str(line.line_id)
                if line.barcode and not queued_invoice["article_ids"].get(line_id):
                    exception_index = next((
                        index for index, exception in enumerate(queued_invoice["exceptions"])
                        if str(exception.get("Línea")) == line_id
                    ), 0)
                    barcode_rows.append({
                        "Factura": queued_invoice["invoice"].invoice_id,
                        "Proveedor": queued_invoice["invoice"].supplier_name or queued_invoice["source"].name,
                        "Línea": line_id,
                        "Producto del XML": line.description,
                        "Código de barras": line.barcode,
                        "Acción requerida": "Vincular este GTIN al producto en EFFI",
                        "_invoice_index": invoice_index,
                        "_exception_index": exception_index,
                        "_source_name": queued_invoice["source"].name,
                    })
        gtin_search = st.text_input("Buscar factura, producto o GTIN", key="gtin_search")
        filtered_barcodes = [
            row for row in barcode_rows
            if not gtin_search.strip() or gtin_search.strip().casefold() in " ".join(
                str(row.get(key, "")) for key in ("Factura", "Proveedor", "Producto del XML", "Código de barras")
            ).casefold()
        ]
        if filtered_barcodes:
            st.dataframe(
                [{key: value for key, value in row.items() if not key.startswith("_")}
                 for row in filtered_barcodes],
                use_container_width=True,
                hide_index=True,
            )
            selected_gtin = st.selectbox(
                "Excepción para revisar",
                range(len(filtered_barcodes)),
                format_func=lambda index: (
                    f'{filtered_barcodes[index]["Factura"]} · {filtered_barcodes[index]["Código de barras"]} · '
                    f'{filtered_barcodes[index]["Producto del XML"]}'
                ),
                key="gtin_review_selection",
            )
            selected_exception = filtered_barcodes[selected_gtin]
            st.button(
                "Abrir factura relacionada",
                on_click=_open_product_review,
                args=(selected_exception["_invoice_index"], selected_exception["_exception_index"], selected_exception["_source_name"]),
            )
        else:
            st.success("No hay códigos de barras pendientes de vincular en las facturas cargadas.")
        return

    if main_view == "Historial":
        st.markdown('<div class="page-eyebrow">Auditoría</div>', unsafe_allow_html=True)
        st.markdown('<h1>Historial de revisión</h1>', unsafe_allow_html=True)
        st.markdown('<p class="page-subtitle">Decisiones de producto, costos y aprobaciones registradas.</p>', unsafe_allow_html=True)
        history = _history_rows(queue, review_state)
        history_search = st.text_input("Buscar factura, proveedor o responsable", key="history_search")
        history_statuses = ["Todos"] + sorted({row["Estado"] for row in history})
        history_status = st.selectbox("Tipo de decisión", history_statuses, key="history_status")
        normalized_search = history_search.strip().casefold()
        filtered_history = [
            row for row in history
            if (history_status == "Todos" or row["Estado"] == history_status)
            and (not normalized_search or normalized_search in " ".join(str(value) for value in row.values()).casefold())
        ]
        if filtered_history:
            st.dataframe(filtered_history, use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay decisiones guardadas para mostrar.")
        return

    invoice_selector_key = "review_invoice_selection"
    pending_invoice_target = st.session_state.pop("next_review_invoice_index", None)
    if pending_invoice_target is not None:
        st.session_state[invoice_selector_key] = pending_invoice_target
    selected_index = min(max(int(st.session_state.get(invoice_selector_key, 0)), 0), len(queue) - 1)
    queue_item = queue[selected_index]
    invoice = queue_item["invoice"]
    source = queue_item["source"]
    line_by_id = {str(line.line_id): line for line in invoice.lines}
    invoice_reviews = review_state.get(source.name, {})

    approval_is_current = _invoice_approval_is_current(source, invoice_reviews)
    invoice_ready = _invoice_review_complete(queue_item, invoice_reviews)
    approver_key = f"approver_{source.name}"
    approver = str(st.session_state.get(approver_key, reviewer_name))

    sync_column, generate_column = st.columns(2, gap="small")
    with sync_column:
        if st.button("⟳  Sincronizar", use_container_width=True, key="sync_queue"):
            try:
                with st.spinner("Generando facturas y asignando productos por GTIN..."):
                    sync_result = _sync_invoice_workbooks(force=True)
                st.session_state["invoice_sync_notice"] = (
                    f'{len(sync_result["generated"])} factura(s) actualizada(s); '
                    f'{sync_result["mapped_barcodes"]} GTIN asignado(s), '
                    f'{sync_result["unmapped_barcodes"]} pendiente(s) de excepción.'
                )
            except Exception as error:
                st.session_state["invoice_sync_error"] = str(error)
            st.rerun()
    with generate_column:
        if st.button(
            "↓  Generar Excel revisado",
            type="primary",
            use_container_width=True,
            disabled=not invoice_ready or (not approval_is_current and not approver.strip()),
            key="generate_approved_workbook",
        ):
            try:
                final_reviews = {
                    line_id: value for line_id, value in invoice_reviews.items()
                    if line_id not in {INVOICE_APPROVAL_KEY, INVOICE_APPROVAL_HISTORY_KEY}
                    and isinstance(value, dict) and (
                        value.get("status") in {"confirmed", "no_match"}
                        or value.get("final_unit_cost") is not None
                        or value.get("is_bonus")
                    )
                }
                _reviewed_workbook(queue_item, final_reviews)
                if not approval_is_current:
                    _record_invoice_approval(review_state, source, invoice_reviews, approver)
                st.rerun()
            except Exception as error:
                st.error(f"No se pudo aprobar la factura ni generar el Excel final: {error}")
    if invoice_ready and not approval_is_current:
        approver = st.text_input("Nombre de quien aprueba", value=reviewer_name, key=approver_key)
    if not invoice_ready:
        st.caption("Resuelve todas las excepciones para habilitar la aprobación final.")
    elif not approval_is_current:
        st.caption("La aprobación registra responsable, hora y versión de los datos revisados.")

    st.markdown(
        '<div class="summary-grid">'
        f'<div class="summary-card"><div class="summary-label">Pendientes de revisión</div>'
        f'<div class="summary-value">{metrics["pending_invoices"]}</div>'
        f'<div class="summary-note">{metrics["pending_products"]} productos por revisar</div></div>'
        f'<div class="summary-card"><div class="summary-label">En proceso</div>'
        f'<div class="summary-value">{metrics["in_progress_invoices"]}</div>'
        f'<div class="summary-note">{metrics["pending_products"]} productos restantes</div></div>'
        f'<div class="summary-card"><div class="summary-label">Excepciones GTIN</div>'
        f'<div class="summary-value">{metrics["gtin_exceptions"]}</div>'
        '<div class="summary-note">Requieren acción en EFFI</div></div>'
        f'<div class="summary-card"><div class="summary-label">Revisadas hoy</div>'
        f'<div class="summary-value">{metrics["reviewed_today"]}</div>'
        '<div class="summary-note">Decisiones y aprobaciones registradas</div></div></div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="section-title">Facturas</div>', unsafe_allow_html=True)
    search_column, status_column, supplier_column, date_column = st.columns([2, 1, 1, 1])
    invoice_search = search_column.text_input("Buscar por factura, proveedor o NIT", key="invoice_search")
    invoice_status = status_column.selectbox(
        "Filtrar por estado",
        ("Todos los estados", "Pendiente", "En revisión", "Pendiente de aprobación", "Revisada"),
        key="invoice_status_filter",
    )
    suppliers = [item["invoice"].supplier_name or item["source"].name for item in queue]
    supplier_filter = supplier_column.selectbox(
        "Filtrar por proveedor", ["Todos los proveedores"] + sorted(set(suppliers)), key="invoice_supplier_filter"
    )
    date_filter = date_column.selectbox(
        "Filtrar por fecha", ("Últimos 30 días", "Hoy", "Esta semana", "Todas las fechas"), key="invoice_date_filter"
    )
    dashboard_rows = []
    for invoice_index, item in enumerate(queue):
        invoice_state = review_state.get(item["source"].name, {})
        try:
            xml_updated = datetime.fromtimestamp(item["source"].stat().st_mtime).strftime("%d %b %Y, %H:%M")
        except OSError:
            xml_updated = "Sin dato"
        dashboard_rows.append({
            "Factura": item["invoice"].invoice_id,
            "Proveedor": item["invoice"].supplier_name or item["source"].name,
            "NIT": item["invoice"].supplier_tax_id or "Sin NIT",
            "Fecha": item["invoice"].issue_date or "Sin fecha",
            "Productos": item["product_count"],
            "Excepciones": len(item["exceptions"]),
            "Estado de revisión": _review_status(item, invoice_state),
            "XML actualizado": xml_updated,
            "_invoice_index": invoice_index,
        })
    today = date.today()
    normalized_search = invoice_search.strip().casefold()
    filtered_invoices = []
    for row in dashboard_rows:
        if invoice_status != "Todos los estados" and row["Estado de revisión"] != invoice_status:
            continue
        if supplier_filter != "Todos los proveedores" and row["Proveedor"] != supplier_filter:
            continue
        if normalized_search and normalized_search not in " ".join(
            str(row[key]) for key in ("Factura", "Proveedor", "NIT")
        ).casefold():
            continue
        try:
            issue_date = date.fromisoformat(row["Fecha"][:10])
        except (TypeError, ValueError):
            issue_date = None
        if date_filter == "Hoy" and issue_date != today:
            continue
        if date_filter == "Esta semana" and (issue_date is None or (today - issue_date).days > 7):
            continue
        if date_filter == "Últimos 30 días" and (issue_date is None or (today - issue_date).days > 30):
            continue
        filtered_invoices.append(row)
    if filtered_invoices:
        st.dataframe(
            [{key: value for key, value in row.items() if not key.startswith("_")}
             for row in filtered_invoices],
            use_container_width=True,
            hide_index=True,
        )
        st.caption(f"Mostrando {len(filtered_invoices)} de {len(dashboard_rows)} facturas")
        visible_indexes = [row["_invoice_index"] for row in filtered_invoices]
        selected_table_index = st.selectbox(
            "Abrir factura",
            range(len(filtered_invoices)),
            index=visible_indexes.index(selected_index) if selected_index in visible_indexes else 0,
            format_func=lambda index: (
                f'{filtered_invoices[index]["Factura"]} · {filtered_invoices[index]["Proveedor"]}'
            ),
            key="invoice_table_selection",
        )
        st.button(
            "Abrir revisión de factura",
            on_click=_select_invoice,
            args=(visible_indexes[selected_table_index],),
        )
    else:
        st.info("No hay facturas que coincidan con estos filtros.")

    product_rows = _invoice_product_rows(queue_item, invoice_reviews)
    pending_count = _pending_review_count(queue_item, invoice_reviews)
    product_row_by_line = {
        line_id: product_row
        for product_row in product_rows
        for line_id in product_row["line_ids"]
    }
    exceptions = queue_item["exceptions"]
    cursor_key = f"review_cursor_{source.name}"
    human_review_indexes = [
        index for index, exception in enumerate(exceptions)
        if _requires_human_decision(exception)
    ]
    if cursor_key not in st.session_state:
        st.session_state[cursor_key] = _next_pending_index(exceptions, invoice_reviews, 0) or 0
    current_index = min(max(int(st.session_state[cursor_key]), 0), max(len(exceptions) - 1, 0))

    st.markdown(f"## Excepciones · Factura {invoice.invoice_id}")
    st.caption(f"Proveedor: {invoice.supplier_name or 'Sin nombre'} · Fecha: {invoice.issue_date or 'Sin fecha'}")
    if not exceptions:
        st.success("Esta factura no tiene excepciones para revisar.")
        item = None
        line_id = ""
        line = None
        review = {}
        review_key = ""
        candidate = None
        confidence = 0
    else:
        current_item = exceptions[current_index]
        current_line_id = str(current_item.get("Línea"))
        review = invoice_reviews.get(current_line_id, {})
        status = review.get("status")
        manual_review_required = _requires_human_decision(current_item)
        status_label = (
            "Confirmado" if status == "confirmed"
            else "Sin coincidencia" if status == "no_match"
            else "Revisión humana" if manual_review_required
            else "Resuelto automáticamente"
        )
        if manual_review_required:
            manual_position = human_review_indexes.index(current_index) + 1
            st.progress(manual_position / len(human_review_indexes))
            st.caption(
                f"Corrección {manual_position} de {len(human_review_indexes)} · "
                f"Línea XML {current_line_id} · {status_label}"
            )
        else:
            st.caption(f"Línea XML {current_line_id} · {status_label}")
        previous_column, next_column = st.columns(2)
        previous_indexes = [index for index in human_review_indexes if index < current_index]
        next_indexes = [index for index in human_review_indexes if index > current_index]
        with previous_column:
            if st.button(
                "Anterior",
                disabled=not previous_indexes,
                key=f"previous_{source.name}",
            ):
                st.session_state[cursor_key] = previous_indexes[-1]
                st.rerun()
        with next_column:
            if st.button(
                "Siguiente",
                disabled=not next_indexes,
                key=f"next_{source.name}",
            ):
                st.session_state[cursor_key] = next_indexes[0]
                st.rerun()
        if pending_count == 0 and not _invoice_approval_is_current(source, invoice_reviews):
            next_invoice = _next_pending_invoice_index(queue, review_state, selected_index + 1)
            if next_invoice is None:
                st.info("Todas las excepciones tienen decisión. Falta la aprobación final de la factura.")
            else:
                st.info("Excepciones revisadas. Puedes aprobar esta factura o continuar con la siguiente.")

        item = current_item
        line_id = str(item.get("Línea"))
        line = line_by_id.get(line_id)
        review_key = f"{source.stem}_{line_id}"
        candidate = _product_from_exception(item)
        if line and line.barcode and not item.get("Artículo EFFI"):
            candidate = None
        confidence = item.get("Puntaje")

        st.markdown(f"#### Producto de la factura · Línea {line_id}")
        xml_column, azure_column = st.columns(2, gap="large")
        with xml_column:
            st.markdown('<div class="compare-panel xml"><span class="eyebrow">Factura · XML</span></div>', unsafe_allow_html=True)
            description = line.description if line else item.get("Descripción factura")
            st.markdown(f"**{html.escape(str(description or 'Descripción vacía'))}**")
            st.caption(f"Código de barras: {_display(line.barcode if line else item.get('Código de barras'))}")
            if line and line.barcode:
                if item.get("Artículo EFFI"):
                    st.success(f"GTIN registrado en EFFI · IdProducto {item['Artículo EFFI']}")
                else:
                    st.error(
                        f"Código de barras {line.barcode} no registrado en EFFI para el producto "
                        f"'{html.escape(line.description or 'sin descripción')}'. "
                        "Vincula este código al producto en EFFI; no se buscará por descripción."
                    )
            invoice_presentation = item.get("Presentación factura")
            if invoice_presentation is None or str(invoice_presentation).strip().lower() == "none":
                invoice_presentation = ", ".join(sorted(_presentation_tokens(description or "")))
            st.caption(f"Presentación detectada: {_format_presentation(invoice_presentation)}")
            amount_columns = st.columns(3)
            amount_columns[0].metric("Cantidad", _display(line.quantity if line else item.get("Cantidad")))
            amount_columns[1].metric("Precio unitario", _money(line.unit_price if line else item.get("Precio unitario XML")))
            amount_columns[2].metric("Descuento", _money(line.discount_total if line else item.get("Descuento XML")))
            st.metric("Neto declarado en factura", _money(line.net_amount if line else item.get("Neto XML")))
            if line and _is_full_discount_bonus(line):
                st.info("Bonificación del 100% detectada automáticamente; confirma a qué producto pagado pertenece.")
            if item.get("Bonificación consolidada"):
                st.info(
                    f"Bonificación consolidada: {item.get('Unidades totales del producto')} unidades "
                    f"de las líneas {item.get('Líneas XML consolidadas')}. "
                    f"Costo efectivo: {_money(item.get('Costo unitario efectivo'))} por unidad; "
                    f"neto agrupado: {_money(item.get('Neto agrupado'))}."
                )
            _render_cost_editor(
                source,
                invoice,
                line,
                invoice_reviews,
                review_state,
                line_by_id,
                product_row_by_line,
                reviewer_name,
            )
        with azure_column:
            st.markdown('<div class="compare-panel"><span class="eyebrow">Sugerencia EFFI</span></div>', unsafe_allow_html=True)
            _render_product(candidate, confidence)
            if review.get("status") == "confirmed":
                st.success(f"Confirmado: EFFI {review.get('product_id')} · {_display(review.get('product_name'))}")
            elif review.get("status") == "no_match":
                st.warning("Marcado como sin coincidencia confiable.")
            elif candidate:
                if manual_review_required:
                    st.caption("La sugerencia del modelo requiere confirmación por baja confianza.")
                    if st.button("Aceptar sugerencia", type="primary", key=f"accept_suggestion_{review_key}"):
                        _record_review(review_state, source, line_id, "confirmed", candidate, reviewer_name)
                        st.rerun()
                else:
                    st.success(f"Coincidencia asignada automáticamente · EFFI {candidate.get('IdProducto')}")

        with st.expander(
            "Consultar código exacto en EFFI" if line and line.barcode
            else "Corregir artículo manualmente" if not manual_review_required
            else "Editar manualmente · buscar en EFFI",
            expanded=manual_review_required and (candidate is None or float(confidence or 0) < 0.72),
        ):
            search_form = st.form(f"search_form_{review_key}")
            with search_form:
                query_text = st.text_input(
                    "Código de barras de la factura" if line and line.barcode
                    else "Producto / IdProducto EFFI",
                    value=line.barcode if line and line.barcode else str(item.get("Descripción factura") or ""),
                    placeholder="Buscar por nombre, referencia o IdProducto",
                    key=f"query_{review_key}",
                    disabled=bool(line and line.barcode),
                )
                search_submitted = st.form_submit_button(
                    "Consultar GTIN" if line and line.barcode else "Buscar en EFFI",
                    type="primary",
                )
            hits_key = f"azure_hits_{review_key}"
            if search_submitted:
                try:
                    with st.spinner("Consultando catálogo EFFI..."):
                        st.session_state[hits_key] = (
                            search_products_by_barcode(line.barcode)
                            if line and line.barcode
                            else search_products(query_text, limit=80)
                        )
                except Exception as error:
                    st.error(f"No se pudo consultar Azure: {error}")

            hits = st.session_state.get(hits_key, [])
            options = hits or ([candidate] if candidate else [])
            if options:
                st.dataframe(
                    [
                        {
                            "IdProducto": product.get("IdProducto"),
                            "Nombre EFFI": product.get("Nombre"),
                            "Referencia": product.get("Referencia"),
                            "Marca": product.get("Marca"),
                            "Categoría": product.get("Categoria"),
                            "Precio normal": _money(product.get("PrecioConImpuestoTarifaNormal")),
                            "Precio mayorista": _money(product.get("PrecioConImpuestoTarifaMayorista")),
                        }
                        for product in options
                    ],
                    use_container_width=True,
                    hide_index=True,
                )
                selected_product_index = st.selectbox(
                    "Producto que vas a confirmar",
                    range(len(options)),
                    format_func=lambda index: _product_label(options[index]),
                    key=f"product_{review_key}",
                )
                selected_product = options[selected_product_index]
                _render_product(selected_product)
                if review.get("status") not in {"confirmed", "no_match"} and st.button(
                    "Confirmar producto", type="primary", key=f"confirm_{review_key}"
                ):
                    _record_review(review_state, source, line_id, "confirmed", selected_product, reviewer_name)
                    st.session_state.pop(hits_key, None)
                    next_pending = _next_pending_index(
                        exceptions, review_state[source.name], current_index + 1
                    )
                    if next_pending is not None:
                        st.session_state[cursor_key] = next_pending
                    else:
                        next_invoice = _next_pending_invoice_index(queue, review_state, selected_index + 1)
                        if next_invoice is not None:
                            st.session_state["next_review_invoice_index"] = next_invoice
                    st.success("Decisión guardada.")
                    st.rerun()
            else:
                if line and line.barcode:
                    st.error(
                        f"El GTIN {line.barcode} no aparece en EFFI. Vincula ese código al producto desde EFFI."
                    )
                else:
                    st.info("Sin resultados todavía. Ajusta la búsqueda y consulta EFFI.")
            if review.get("status") in {"confirmed", "no_match"}:
                if st.button("Reabrir esta línea", key=f"reopen_{review_key}"):
                    _record_review(review_state, source, line_id, "pending", reviewed_by=reviewer_name)
                    st.session_state[cursor_key] = current_index
                    st.rerun()
            elif st.button("No coincide", key=f"reject_{review_key}"):
                _record_review(review_state, source, line_id, "no_match", reviewed_by=reviewer_name)
                st.session_state.pop(hits_key, None)
                next_pending = _next_pending_index(
                    exceptions, review_state[source.name], current_index + 1
                )
                if next_pending is not None:
                    st.session_state[cursor_key] = next_pending
                else:
                    next_invoice = _next_pending_invoice_index(queue, review_state, selected_index + 1)
                    if next_invoice is not None:
                        st.session_state["next_review_invoice_index"] = next_invoice
                st.info("Se guardó como excepción sin coincidencia.")
                st.rerun()

    st.markdown("---")
    base_output = queue_item["output"] or (
        DEFAULT_OUTPUT_DIR / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"
    )
    reviewed_path = base_output.with_name(base_output.stem.removesuffix("_effi") + "_revisado.xlsx")
    if approval_is_current and reviewed_path.is_file():
        st.success(
            f"Aprobada por {invoice_reviews[INVOICE_APPROVAL_KEY].get('approved_by', 'Sin nombre')} · "
            f"{reviewed_path.name}"
        )
        st.download_button(
            "Descargar Excel aprobado",
            data=reviewed_path.read_bytes(),
            file_name=reviewed_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()