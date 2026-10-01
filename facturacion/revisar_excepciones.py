from __future__ import annotations

import html
import json
import os
import sys
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
from typing import Any

import streamlit as st
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
REPOSITORY_ROOT = ROOT.parent
if str(REPOSITORY_ROOT) not in sys.path:
    sys.path.insert(0, str(REPOSITORY_ROOT))

from facturacion.azure_codes import search_products
from facturacion.generar_importacion import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_XML_DIR,
    _presentation_tokens,
    _safe_stem,
    parse_invoice,
)


REVIEW_FILE = DEFAULT_OUTPUT_DIR / "revisiones_manuales.json"
DETAIL_SHEET = "Hoja1"


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
    queue = []
    for source in sorted(DEFAULT_XML_DIR.glob("*.xml")):
        try:
            invoice = parse_invoice(source)
        except (OSError, ValueError):
            continue
        output = DEFAULT_OUTPUT_DIR / f"{_safe_stem(source, invoice.invoice_id)}_effi.xlsx"
        if not output.is_file():
            continue
        exceptions = _read_exception_rows(output)
        if exceptions:
            queue.append({
                "source": source,
                "invoice": invoice,
                "output": output,
                "exceptions": exceptions,
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
    state: dict[str, Any], source: Path, line_id: str, status: str, product: dict[str, Any] | None = None
) -> None:
    invoice_reviews = state.setdefault(source.name, {})
    invoice_reviews[line_id] = {
        "status": status,
        "product_id": product.get("IdProducto") if product else None,
        "product_name": product.get("Nombre") if product else None,
        "reference": product.get("Referencia") if product else None,
        "reviewed_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    _save_review_state(state)


def _next_pending_index(
    exceptions: list[dict[str, Any]],
    invoice_reviews: dict[str, Any],
    start_index: int,
) -> int | None:
    if not exceptions:
        return None
    indexes = list(range(start_index, len(exceptions))) + list(range(0, min(start_index, len(exceptions))))
    for index in indexes:
        line_id = str(exceptions[index].get("Línea"))
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
    workbook = load_workbook(original)
    sheet = workbook[DETAIL_SHEET] if DETAIL_SHEET in workbook.sheetnames else workbook.active
    line_rows = {str(line.line_id): row for row, line in enumerate(invoice.lines, start=2)}
    for line_id, review in approved.items():
        row = line_rows.get(str(line_id))
        if row is None:
            continue
        if review.get("status") == "confirmed" and review.get("product_id") is not None:
            sheet.cell(row=row, column=1, value=review["product_id"])
        elif review.get("status") == "no_match":
            sheet.cell(row=row, column=1).value = None
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
      :root { --ink: #172a2a; --paper: #f3f4ef; --line: #d8ded8; --teal: #087f78; --coral: #c45135; }
      .stApp { background: var(--paper); color: var(--ink); }
      [data-testid="stSidebar"] { background: #172a2a; }
      [data-testid="stSidebar"] * { color: #f4f5ef; }
            [data-testid="stSidebar"] [data-testid="stRadio"] label {
                align-items: flex-start;
                border-bottom: 1px solid rgba(244,245,239,.18);
                padding: .62rem .25rem;
            }
            [data-testid="stSidebar"] [data-testid="stRadio"] label p {
                color: #f4f5ef !important;
                line-height: 1.3;
                overflow-wrap: anywhere;
                white-space: normal !important;
            }
            [data-testid="stMetricValue"] {
                font-size: 1.08rem;
                overflow-wrap: anywhere;
                white-space: normal;
            }
      h1, h2, h3 { color: var(--ink); font-family: Georgia, "Times New Roman", serif; }
      h1 { font-size: 2.15rem; }
      .eyebrow { color: var(--coral); font-size: .76rem; font-weight: 700; text-transform: uppercase; }
      .compare-panel { border-top: 3px solid var(--teal); padding: 1rem 0 .25rem; min-height: 7rem; }
      .compare-panel.xml { border-color: var(--coral); }
      div.stButton > button[kind="primary"] { background: var(--teal); border-color: var(--teal); }
      div.stButton > button { border-radius: 4px; }
      [data-testid="stMetric"] { background: transparent; border-bottom: 1px solid var(--line); padding: .35rem 0; }
      div[data-testid="stDataFrame"] { border: 1px solid var(--line); }
    </style>
    """, unsafe_allow_html=True)


def main() -> None:
    st.set_page_config(
        page_title="Revisión EFFI · HITL",
        page_icon="E",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    _render_styles()
    st.markdown('<div class="eyebrow">Facturación · Human in the loop</div>', unsafe_allow_html=True)
    st.title("Revisión de excepciones")
    st.caption("Compara el detalle de factura con el catálogo EFFI y confirma cada artículo manualmente.")

    queue = _invoice_queue()
    if not queue:
        st.info("No hay facturas con excepciones. Genera primero los libros con --azure-lookup.")
        return

    review_state = _load_review_state()
    invoice_labels = []
    for item in queue:
        invoice_review = review_state.get(item["source"].name, {})
        pending = sum(
            invoice_review.get(str(exception.get("Línea")), {}).get("status") not in {"confirmed", "no_match"}
            for exception in item["exceptions"]
        )
        invoice_labels.append(
            f"{item['invoice'].invoice_id} · {item['invoice'].supplier_name or item['source'].name} · "
            f"{pending} pendientes"
        )
    invoice_selector_key = "review_invoice_selection"
    pending_invoice_target = st.session_state.pop("next_review_invoice_index", None)
    if pending_invoice_target is not None:
        st.session_state[invoice_selector_key] = pending_invoice_target
    st.sidebar.markdown("### Facturas")
    st.sidebar.caption("Revisión secuencial por factura")
    selected_index = st.sidebar.radio(
        "Facturas pendientes",
        range(len(queue)),
        format_func=lambda index: invoice_labels[index],
        key=invoice_selector_key,
        label_visibility="collapsed",
    )
    queue_item = queue[selected_index]
    invoice = queue_item["invoice"]
    source = queue_item["source"]
    line_by_id = {str(line.line_id): line for line in invoice.lines}
    invoice_reviews = review_state.get(source.name, {})

    st.sidebar.markdown("---")
    st.sidebar.markdown(f"**{invoice.supplier_name or 'Proveedor sin nombre'}**")
    st.sidebar.caption(f"Factura {invoice.invoice_id} · {invoice.issue_date or 'Fecha no informada'}")
    confirmed_count = sum(review.get("status") == "confirmed" for review in invoice_reviews.values())
    rejected_count = sum(review.get("status") == "no_match" for review in invoice_reviews.values())
    pending_count = max(len(queue_item["exceptions"]) - confirmed_count - rejected_count, 0)
    metric_a, metric_b, metric_c = st.columns(3)
    metric_a.metric("Pendientes", pending_count)
    metric_b.metric("Confirmados", confirmed_count)
    metric_c.metric("Sin coincidencia", rejected_count)

    exceptions = queue_item["exceptions"]
    cursor_key = f"review_cursor_{source.name}"
    if cursor_key not in st.session_state:
        st.session_state[cursor_key] = _next_pending_index(exceptions, invoice_reviews, 0) or 0
    current_index = min(max(int(st.session_state[cursor_key]), 0), max(len(exceptions) - 1, 0))

    st.markdown(f"### Factura {invoice.invoice_id}")
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
        status_label = "Confirmado" if status == "confirmed" else "Sin coincidencia" if status == "no_match" else "Pendiente"
        st.progress((current_index + 1) / len(exceptions))
        st.caption(f"Excepción {current_index + 1} de {len(exceptions)} · Línea XML {current_line_id} · {status_label}")
        previous_column, next_column = st.columns(2)
        with previous_column:
            if st.button("Anterior", disabled=current_index == 0, key=f"previous_{source.name}"):
                st.session_state[cursor_key] = current_index - 1
                st.rerun()
        with next_column:
            if st.button("Siguiente", disabled=current_index >= len(exceptions) - 1, key=f"next_{source.name}"):
                st.session_state[cursor_key] = current_index + 1
                st.rerun()
        if pending_count == 0:
            next_invoice = _next_pending_invoice_index(queue, review_state, selected_index + 1)
            if next_invoice is None:
                st.success("Todas las excepciones de todas las facturas ya tienen una decisión.")
            else:
                st.success("Factura revisada. La próxima factura pendiente se abrirá al terminar esta línea.")

        item = current_item
        line_id = str(item.get("Línea"))
        line = line_by_id.get(line_id)
        review_key = f"{source.stem}_{line_id}"
        candidate = _product_from_exception(item)
        confidence = item.get("Puntaje")

        st.markdown(f"#### Producto de la factura · Línea {line_id}")
        xml_column, azure_column = st.columns(2, gap="large")
        with xml_column:
            st.markdown('<div class="compare-panel xml"><span class="eyebrow">Factura · XML</span></div>', unsafe_allow_html=True)
            description = line.description if line else item.get("Descripción factura")
            st.markdown(f"**{html.escape(str(description or 'Descripción vacía'))}**")
            st.caption(f"Código de barras: {_display(line.barcode if line else item.get('Código de barras'))}")
            invoice_presentation = item.get("Presentación factura")
            if invoice_presentation is None or str(invoice_presentation).strip().lower() == "none":
                invoice_presentation = ", ".join(sorted(_presentation_tokens(description or "")))
            st.caption(f"Presentación detectada: {_format_presentation(invoice_presentation)}")
            amount_columns = st.columns(3)
            amount_columns[0].metric("Cantidad", _display(line.quantity if line else item.get("Cantidad")))
            amount_columns[1].metric("Precio unitario", _money(line.unit_price if line else item.get("Precio unitario XML")))
            amount_columns[2].metric("Descuento", _money(line.discount_total if line else item.get("Descuento XML")))
            st.metric("Neto declarado en factura", _money(line.net_amount if line else item.get("Neto XML")))
        with azure_column:
            st.markdown('<div class="compare-panel"><span class="eyebrow">EFFI · producto sugerido</span></div>', unsafe_allow_html=True)
            _render_product(candidate, confidence)
            if review.get("status") == "confirmed":
                st.success(f"Confirmado: EFFI {review.get('product_id')} · {_display(review.get('product_name'))}")
            elif review.get("status") == "no_match":
                st.warning("Marcado como sin coincidencia confiable.")

        with st.expander(
            "Buscar manualmente en EFFI",
            expanded=candidate is None or float(confidence or 0) < 0.72,
        ):
            search_form = st.form(f"search_form_{review_key}")
            with search_form:
                query_text = st.text_input(
                    "Buscar por descripción, marca, referencia o código",
                    value=str(item.get("Descripción factura") or ""),
                    key=f"query_{review_key}",
                )
                search_submitted = st.form_submit_button("Buscar en EFFI", type="primary")
            hits_key = f"azure_hits_{review_key}"
            if search_submitted:
                try:
                    with st.spinner("Consultando catálogo EFFI..."):
                        st.session_state[hits_key] = search_products(query_text, limit=80)
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
                    "Confirmar artículo", type="primary", key=f"confirm_{review_key}"
                ):
                    _record_review(review_state, source, line_id, "confirmed", selected_product)
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
                st.info("Sin resultados todavía. Ajusta la búsqueda y consulta Azure.")
            if review.get("status") in {"confirmed", "no_match"}:
                if st.button("Reabrir esta línea", key=f"reopen_{review_key}"):
                    _record_review(review_state, source, line_id, "pending")
                    st.session_state[cursor_key] = current_index
                    st.rerun()
            elif st.button("No coincide", key=f"reject_{review_key}"):
                _record_review(review_state, source, line_id, "no_match")
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

    final_reviews = {
        line_id: value for line_id, value in invoice_reviews.items()
        if value.get("status") == "no_match"
        or (value.get("status") == "confirmed" and value.get("product_id") is not None)
    }
    st.markdown("---")
    if final_reviews:
        reviewed_path = queue_item["output"].with_name(
            queue_item["output"].stem.removesuffix("_effi") + "_revisado.xlsx"
        )
        if st.button("Generar Excel revisado", type="primary"):
            try:
                _reviewed_workbook(queue_item, final_reviews)
            except Exception as error:
                st.error(f"No se pudo generar el Excel revisado: {error}")
        if reviewed_path.is_file():
            st.success(f"Excel revisado de la factura {invoice.invoice_id}: {reviewed_path.name}")
            st.download_button(
                "Descargar Excel revisado",
                data=reviewed_path.read_bytes(),
                file_name=reviewed_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.caption("Decide al menos una excepción para generar el Excel revisado.")


if __name__ == "__main__":
    main()