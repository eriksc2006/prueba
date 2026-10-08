from __future__ import annotations

import tempfile
import unittest
from contextlib import nullcontext
from decimal import Decimal
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from facturacion.generar_importacion import (
    DEFAULT_TAX_CODES,
    InvoiceParseError,
    generate_workbook,
    parse_invoice,
)


ROOT = Path(__file__).resolve().parents[2]
XML_DIR = ROOT / "facturacion" / "files" / "xml"
TEMPLATE = ROOT / "facturacion" / "files" / "template" / "plantilla_importacion_conceptos_compra.xlsx"


class InvoiceParsingTests(unittest.TestCase):
    def test_confident_no_barcode_match_uses_variant_brand_and_margin(self) -> None:
        from facturacion.generar_importacion import _confident_no_barcode_match
        from facturacion.revisar_excepciones import _build_invoice_queue
        from facturacion.revisar_excepciones import _requires_human_decision

        queue = _build_invoice_queue()
        invoice = next(item["invoice"] for item in queue if item["invoice"].invoice_id == "FECG315045")
        line = next(line for line in invoice.lines if "12.013" in line.description)
        candidate = {
            "IdProducto": 5546,
            "Nombre": "TNT. NOUVELLE #12.013",
            "Referencia": "005563",
            "Marca": "EURO STYLE",
        }

        self.assertTrue(_confident_no_barcode_match(line, candidate, 0.11, invoice.supplier_name))
        self.assertFalse(_confident_no_barcode_match(
            line, {**candidate, "Nombre": "TNT. NOUVELLE #10.013"}, 0.11, invoice.supplier_name
        ))
        self.assertFalse(_confident_no_barcode_match(line, candidate, 0.05, invoice.supplier_name))
        self.assertFalse(_requires_human_decision({
            "_automatic_match": True,
            "Tipo": "Candidato cercano; validar identidad",
        }))
        self.assertTrue(_requires_human_decision({
            "_automatic_match": True,
            "Tipo": "Candidato cercano; validar identidad; Bonificación 100% detectada",
        }))

    def test_f2_tax_calculation_matches_effis_example(self) -> None:
        from facturacion.revisar_excepciones import _price_including_taxes

        self.assertEqual(
            _price_including_taxes(Decimal("4705.88"), (Decimal("19"),)),
            Decimal("5600.00"),
        )

    def test_connection_string_includes_credentials_for_active_directory_password(self) -> None:
        from facturacion.azure_codes import _connection_string

        with patch.dict("os.environ", {
            "SQL_SERVER": "server.example",
            "SQL_DATABASE": "database",
            "SQL_USERNAME": "user@example.com",
            "SQL_PASSWORD": "secret",
            "SQL_AUTHENTICATION": "ActiveDirectoryPassword",
        }, clear=False):
            connection_string = _connection_string()

        self.assertIn("Authentication=ActiveDirectoryPassword;", connection_string)
        self.assertIn("UID=user@example.com;", connection_string)
        self.assertIn("PWD=secret;", connection_string)

    def test_connection_string_requires_credentials_for_active_directory_password(self) -> None:
        from facturacion.azure_codes import _connection_string

        with patch.dict("os.environ", {
            "SQL_SERVER": "server.example",
            "SQL_DATABASE": "database",
            "SQL_USERNAME": "",
            "SQL_PASSWORD": "",
            "SQL_AUTHENTICATION": "ActiveDirectoryPassword",
        }, clear=False):
            with self.assertRaisesRegex(RuntimeError, "SQL_USERNAME.*SQL_PASSWORD"):
                _connection_string()

    def test_cost_editor_only_targets_human_exceptions_and_saved_costs(self) -> None:
        from facturacion.revisar_excepciones import _product_cost_editor_line_ids

        product_row = {"line_ids": ["1", "2", "3", "4", "5"]}
        exceptions_by_line = {
            "1": {"Tipo": "Código de barras no registrado en EFFI"},
            "2": {"Tipo": "Bonificación/prorrateo: costo efectivo calculado"},
        }
        invoice_reviews = {
            "3": {"is_bonus": True},
            "4": {"final_unit_cost": "1200.00"},
        }

        self.assertEqual(
            _product_cost_editor_line_ids(product_row, exceptions_by_line, invoice_reviews),
            ["1", "3", "4"],
        )

    def test_parses_attached_document_and_embedded_invoice(self) -> None:
        invoice = parse_invoice(XML_DIR / "ad0811042361016260008d748.xml")
        self.assertEqual(invoice.invoice_id, "FI579400")
        self.assertEqual(invoice.supplier_tax_id, "811042361")
        self.assertEqual(len(invoice.lines), 6)
        self.assertEqual(invoice.lines[0].item_code, "9007009")
        self.assertEqual(invoice.lines[0].quantity, Decimal("12.00"))
        self.assertEqual(invoice.lines[0].unit_price, Decimal("16326.00"))

    def test_parses_buyer_code_and_free_goods(self) -> None:
        invoice = parse_invoice(XML_DIR / "ad089030295502726356720260928112521401.xml")
        free_line = invoice.lines[0]
        self.assertEqual(free_line.item_code, "703931")
        self.assertEqual(free_line.barcode, "7702113703931")
        self.assertEqual(free_line.quantity, Decimal("210.00"))
        self.assertEqual(free_line.unit_price, Decimal("0.00"))
        self.assertEqual(free_line.net_amount, Decimal("0.00"))
        promotional_line = invoice.lines[2]
        self.assertEqual(promotional_line.unit_price, Decimal("18240.00"))
        self.assertEqual(promotional_line.discount_total, Decimal("273600.00"))
        self.assertEqual(
            promotional_line.unit_price * promotional_line.quantity - promotional_line.discount_total,
            promotional_line.net_amount,
        )

    def test_does_not_treat_supplier_item_reference_as_barcode(self) -> None:
        invoice = parse_invoice(XML_DIR / "52089_e617b4bf4e04d5380af089e90fb2f16b17f43f96e25e3d78623dfd2b55c8ba9f70b5b06ff39c2fe285bc58524803dd5b.xml")
        self.assertEqual(invoice.lines[0].item_code, "NOUBAB22")
        self.assertEqual(invoice.lines[0].barcode, "")

    def test_parses_line_discount(self) -> None:
        invoice = parse_invoice(XML_DIR / "ad090136037002826FEMJ00003856.xml")
        line = invoice.lines[0]
        self.assertEqual(line.item_code, "MJPT-235")
        self.assertEqual(line.discount_total, Decimal("11659.50"))
        self.assertEqual(line.tax_rates, (Decimal("19.00"),))

    def test_detects_zero_net_full_discount_bonus(self) -> None:
        from facturacion.generar_importacion import _is_full_discount_bonus
        from facturacion.revisar_excepciones import _requires_human_decision

        invoice = parse_invoice(XML_DIR / "ad089030295502726356720260928112521401.xml")
        bonus_line = next(line for line in invoice.lines if line.net_amount == 0 and line.discount_total > 0)
        paid_line = next(line for line in invoice.lines if line.net_amount > 0)

        self.assertTrue(_is_full_discount_bonus(bonus_line))
        self.assertFalse(_is_full_discount_bonus(paid_line))

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "full_discount.xlsx"
            generate_workbook(
                XML_DIR / "ad089030295502726356720260928112521401.xml",
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={bonus_line.barcode: 98765},
                product_ids_by_description={},
            )
            workbook = load_workbook(output, data_only=True)
            exceptions = workbook["Excepciones"]
            bonus_reasons = [
                exceptions.cell(row, 2).value
                for row in range(2, exceptions.max_row + 1)
                if str(exceptions.cell(row, 1).value) == bonus_line.line_id
            ]
            self.assertTrue(any("Bonificación 100% detectada" in str(reason) for reason in bonus_reasons))
            workbook.close()

        grouped_lines = [line for line in invoice.lines if line.barcode == bonus_line.barcode]
        if len(grouped_lines) > 1:
            self.assertFalse(_requires_human_decision({"Tipo": bonus_reasons[0]}))

    def test_barcode_lookup_leaves_catalog_duplicates_for_human_review(self) -> None:
        from facturacion.azure_codes import get_product_ids_by_barcodes

        cursor = Mock()
        cursor.fetchall.return_value = [
            ("7700000000001", 100),
            ("7700000000001", 101),
            ("7700000000002", 202),
            ("7700000000002", 202),
        ]
        connection = Mock()
        connection.cursor.return_value = cursor
        with patch("facturacion.azure_codes._connect", return_value=connection):
            products = get_product_ids_by_barcodes(("7700000000001", "7700000000002"))

        self.assertEqual(products, {"7700000000002": 202})

    def test_description_lookup_filters_candidates_in_azure(self) -> None:
        from facturacion.azure_codes import get_product_candidates_by_descriptions

        cursor = Mock()
        cursor.execute.side_effect = lambda *_: setattr(cursor, "description", [
            ("IdProducto",), ("CodigoBarras",), ("Nombre",), ("Referencia",),
            ("Categoria",), ("Marca",), ("PrecioConImpuestoTarifaNormal",),
            ("PrecioConImpuestoTarifaMayorista",),
        ])
        cursor.fetchall.return_value = [
            (10, "7701234567890", "SAVITAL PROTEINAS", "SAVITAL", "CABELLO", "SAVITAL", 12000, 11000),
        ]
        connection = Mock()
        connection.cursor.return_value = cursor

        with patch("facturacion.azure_codes._connect", return_value=connection):
            candidates = get_product_candidates_by_descriptions(["SAVITAL PROTEINAS"])

        query, parameters = cursor.execute.call_args.args
        self.assertIn("WHERE IdProducto IS NOT NULL AND", query)
        self.assertIn("Nombre LIKE ?", query)
        self.assertEqual(len(parameters), 8)
        self.assertEqual(candidates["savital proteinas"][0]["IdProducto"], 10)

    def test_description_id_lookup_preserves_full_product_candidates(self) -> None:
        from facturacion.azure_codes import get_product_ids_by_descriptions

        product = {"IdProducto": 10, "Nombre": "SAVITAL PROTEINAS", "Marca": "SAVITAL"}
        with patch(
            "facturacion.azure_codes.get_product_candidates_by_descriptions",
            return_value={"savital proteinas": [product]},
        ) as lookup:
            candidates = get_product_ids_by_descriptions(["SAVITAL PROTEINAS"])

        lookup.assert_called_once_with(["SAVITAL PROTEINAS"])
        self.assertEqual(candidates, {"savital proteinas": [product]})

    def test_manual_description_uses_text_search_when_not_equal_to_invoice_gtin(self) -> None:
        from facturacion.revisar_excepciones import _search_products_for_review

        with patch("facturacion.revisar_excepciones.search_products_by_barcode") as barcode_lookup, patch(
            "facturacion.revisar_excepciones.search_products"
        ) as text_lookup:
            _search_products_for_review("7701234567890", "SAVITAL PROTEINAS")
            barcode_lookup.assert_not_called()
            text_lookup.assert_called_once_with("SAVITAL PROTEINAS", limit=80)

            barcode_lookup.reset_mock()
            text_lookup.reset_mock()
            _search_products_for_review("7701234567890", "7701234567890")
            barcode_lookup.assert_called_once_with("7701234567890")
            text_lookup.assert_not_called()

    def test_singular_description_lookup_rejects_ambiguous_results(self) -> None:
        from facturacion.revisar_excepciones import get_product_by_description

        products = [{"IdProducto": 10}, {"IdProducto": 11}]
        with patch(
            "facturacion.azure_codes.get_product_candidates_by_descriptions",
            return_value={"savital proteinas": products},
        ):
            product = get_product_by_description("SAVITAL PROTEINAS")

        self.assertIsNone(product)

    def test_singular_description_lookup_rejects_single_partial_candidate(self) -> None:
        from facturacion.azure_codes import get_product_by_description

        with patch(
            "facturacion.azure_codes.get_product_candidates_by_descriptions",
            return_value={"savital proteinas": [{"IdProducto": 10, "Nombre": "SAVITAL"}]},
        ):
            product = get_product_by_description("SAVITAL PROTEINAS")

        self.assertIsNone(product)

    def test_rejects_xml_without_invoice(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad.xml"
            path.write_text("<Root />", encoding="utf-8")
            with self.assertRaises(InvoiceParseError):
                parse_invoice(path)

    def test_unresolved_no_barcode_does_not_fallback_to_item_code(self) -> None:
        from facturacion.generar_importacion import _resolved_product

        line = Mock()
        line.barcode = ""
        line.description = "Producto sin coincidencia segura"
        line.item_code = "ABC-999"

        resolved_id, candidate, score, margin, accepted = _resolved_product(
            line,
            None,
            {},
            supplier_name="EURO STYLE",
        )

        self.assertIsNone(resolved_id)
        self.assertFalse(accepted)
        self.assertEqual(candidate, None)
        self.assertEqual(score, 0.0)
        self.assertEqual(margin, 0.0)

    def test_description_id_without_candidate_data_is_not_auto_matched(self) -> None:
        from facturacion.generar_importacion import _resolved_product, _normalize_description

        line = Mock()
        line.barcode = ""
        line.description = "Producto parcial"
        line.item_code = ""

        resolved_id, candidate, score, margin, accepted = _resolved_product(
            line,
            None,
            {_normalize_description(line.description): 5001},
        )

        self.assertIsNone(resolved_id)
        self.assertIsNone(candidate)
        self.assertEqual(score, 0.0)
        self.assertEqual(margin, 0.0)
        self.assertFalse(accepted)

    def test_main_auto_syncs_on_first_load(self) -> None:
        from facturacion import revisar_excepciones

        fake_streamlit = Mock()
        fake_streamlit.session_state = {}
        fake_streamlit.sidebar = Mock()
        fake_streamlit.columns = Mock(return_value=[Mock(), Mock()])
        fake_streamlit.empty = Mock()
        fake_streamlit.container = Mock(return_value=Mock())
        fake_streamlit.dataframe = Mock()
        fake_streamlit.success = Mock()
        fake_streamlit.warning = Mock()
        fake_streamlit.info = Mock()
        fake_streamlit.caption = Mock()
        fake_streamlit.markdown = Mock()
        fake_streamlit.button = Mock(return_value=False)

        fake_streamlit.spinner = Mock(return_value=nullcontext())

        with (
            patch.object(revisar_excepciones, "st", fake_streamlit),
            patch.object(revisar_excepciones, "_render_styles"),
            patch.object(revisar_excepciones, "_invoice_queue", return_value=[]),
            patch.object(revisar_excepciones, "_sync_invoice_workbooks", return_value={"generated": [], "mapped_barcodes": 0, "unmapped_barcodes": 0}) as sync_mock,
        ):
            revisar_excepciones.main()

        sync_mock.assert_called_once_with(force=True)


class WorkbookGenerationTests(unittest.TestCase):
    def test_completed_workbook_download_builds_missing_base_workbook(self) -> None:
        from io import BytesIO

        from facturacion.revisar_excepciones import _build_invoice_queue, _draft_workbook_data

        queue_item = next(item for item in _build_invoice_queue() if item["invoice"].invoice_id == "MAXI610")
        queue_item["output"] = None

        completed_data = _draft_workbook_data(queue_item, {})

        self.assertIsNotNone(completed_data)
        workbook = load_workbook(BytesIO(completed_data), data_only=True)
        self.assertIn("Hoja1", workbook.sheetnames)
        self.assertIn("Excepciones", workbook.sheetnames)
        workbook.close()

    def test_draft_download_applies_safe_match_without_approval(self) -> None:
        from io import BytesIO

        from facturacion.revisar_excepciones import _build_invoice_queue, _draft_workbook_data

        queue = _build_invoice_queue()
        item = next(item for item in queue if item["invoice"].invoice_id == "FECG315045")
        draft_data = _draft_workbook_data(item, {})

        self.assertIsNotNone(draft_data)
        workbook = load_workbook(BytesIO(draft_data), data_only=True)
        self.assertNotIn("Coincidencias automáticas", workbook.sheetnames)

        detail = workbook["Hoja1"]
        self.assertIsNone(detail.cell(row=2, column=1).value)
        workbook.close()

    def test_draft_download_uses_only_current_invoice_and_unsaved_bonus_link(self) -> None:
        from io import BytesIO

        from facturacion import revisar_excepciones

        queue_item = next(
            item for item in revisar_excepciones._build_invoice_queue()
            if item["invoice"].invoice_id == "FECG315045"
        )
        paid_line = next(line for line in queue_item["invoice"].lines if line.item_code == "NOUSHAPROT" and line.net_amount > 0)
        bonus_line = next(line for line in queue_item["invoice"].lines if line.item_code == "NOUSHAPROT" and line.net_amount == 0)
        session_state = {
            f"is_bonus_{queue_item['source'].name}_{bonus_line.line_id}": True,
            f"bonus_link_{queue_item['source'].name}_{bonus_line.line_id}": paid_line.line_id,
        }

        with patch.object(revisar_excepciones.st, "session_state", session_state):
            draft_data = revisar_excepciones._draft_workbook_data(queue_item, {})

        workbook = load_workbook(BytesIO(draft_data), data_only=True)
        detail = workbook["Hoja1"]
        import_row = queue_item["import_row_by_line"][str(paid_line.line_id)]
        self.assertEqual(detail.cell(import_row, 6).value, paid_line.quantity + bonus_line.quantity)
        self.assertEqual(
            Decimal(str(detail.cell(import_row, 7).value)),
            (paid_line.net_amount + bonus_line.net_amount) / (paid_line.quantity + bonus_line.quantity),
        )
        populated_rows = [
            row for row in range(2, detail.max_row + 1)
            if detail.cell(row, 6).value is not None
        ]
        self.assertEqual(len(populated_rows), queue_item["product_count"])
        self.assertTrue(all(
            detail.cell(row, 2).value == 1
            for row in range(2, detail.max_row + 1)
            if detail.cell(row, 6).value is not None
        ))
        workbook.close()

    def test_reviewer_queue_lists_all_invoices_and_consolidated_products(self) -> None:
        from facturacion.revisar_excepciones import _invoice_product_rows, _invoice_queue

        queue = _invoice_queue()
        self.assertEqual(len(queue), 8)
        self.assertEqual(sum(len(item["invoice"].lines) for item in queue), 170)
        self.assertTrue(all(item["product_count"] > 0 for item in queue))
        invoice = next(item for item in queue if item["invoice"].invoice_id == "3166244")
        products = _invoice_product_rows(invoice, {})
        self.assertEqual(len(products), invoice["product_count"])
        bonus = next(product for product in products if product["barcode"] == "7702113048636")
        self.assertEqual(bonus["line_ids"], ["2", "3"])
        self.assertEqual(bonus["quantity"], Decimal("27.00"))

    def test_reviewed_line_counts_as_complete_without_product_assignment(self) -> None:
        from facturacion.revisar_excepciones import _invoice_review_complete

        queue_item = {"exceptions": [{"Línea": "1"}, {"Línea": "2"}]}
        invoice_reviews = {"1": {"status": "reviewed"}, "2": {"status": "confirmed", "product_id": 123}}

        self.assertTrue(_invoice_review_complete(queue_item, invoice_reviews))

    def test_sync_generates_missing_invoice_offline_as_pending_review(self) -> None:
        from facturacion import revisar_excepciones

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            output_dir.mkdir()
            (xml_dir / source.name).write_bytes(source.read_bytes())
            with (
                patch.object(revisar_excepciones, "DEFAULT_XML_DIR", xml_dir),
                patch.object(revisar_excepciones, "DEFAULT_OUTPUT_DIR", output_dir),
                patch.object(revisar_excepciones, "REVIEW_FILE", output_dir / "reviews.json"),
                patch.object(
                    revisar_excepciones,
                    "_cached_barcode_product_ids",
                    side_effect=RuntimeError("Azure firewall blocked the connection"),
                ),
            ):
                result = revisar_excepciones._sync_invoice_workbooks(force=True)
                self.assertIn("Azure firewall", result["lookup_error"])
                self.assertEqual(result["generated"], [source.name])
                queue = revisar_excepciones._build_invoice_queue()

            self.assertEqual(len(queue), 1)
            self.assertFalse(revisar_excepciones._invoice_review_complete(queue[0], {}))
            self.assertTrue(queue[0]["output"].is_file())

    def test_invoice_queue_reuses_cache_until_xml_or_workbook_changes(self) -> None:
        from facturacion import revisar_excepciones

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            output_dir.mkdir()
            (xml_dir / source.name).write_bytes(source.read_bytes())
            with (
                patch.object(revisar_excepciones, "DEFAULT_XML_DIR", xml_dir),
                patch.object(revisar_excepciones, "DEFAULT_OUTPUT_DIR", output_dir),
                patch.object(revisar_excepciones, "REVIEW_FILE", output_dir / "reviews.json"),
                patch.object(revisar_excepciones, "parse_invoice", wraps=parse_invoice) as parse_spy,
            ):
                first_queue = revisar_excepciones._invoice_queue()
                second_queue = revisar_excepciones._invoice_queue()

            self.assertEqual(len(first_queue), 1)
            self.assertEqual(len(second_queue), 1)
            self.assertEqual(parse_spy.call_count, 1)

    def test_opening_dashboard_does_not_forced_sync_azure(self) -> None:
        from facturacion import revisar_excepciones

        with patch.object(revisar_excepciones, "_invoice_queue", return_value=[]), \
                patch.object(revisar_excepciones.st, "session_state", {}), \
                patch.object(revisar_excepciones.st, "markdown"), \
                patch.object(revisar_excepciones.st, "info"), \
                patch.object(revisar_excepciones, "_sync_invoice_workbooks") as sync:
            revisar_excepciones.main()

        sync.assert_not_called()

    def test_invoice_sync_generates_excel_with_azure_barcode_ids(self) -> None:
        from facturacion.revisar_excepciones import _invoice_queue, _sync_invoice_workbooks

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        invoice = parse_invoice(source)
        barcode_ids = {line.barcode: 98765 for line in invoice.lines if line.barcode}
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            (xml_dir / source.name).write_bytes(source.read_bytes())
            with (
                patch("facturacion.revisar_excepciones.DEFAULT_XML_DIR", xml_dir),
                patch("facturacion.revisar_excepciones.DEFAULT_OUTPUT_DIR", output_dir),
                patch("facturacion.revisar_excepciones.REVIEW_FILE", output_dir / "reviews.json"),
                patch("facturacion.azure_codes.get_product_ids_by_barcodes", return_value=barcode_ids),
                patch(
                    "facturacion.revisar_excepciones.get_product_ids_by_descriptions",
                    return_value={},
                ),
            ):
                _invoice_queue()
                sync_result = _sync_invoice_workbooks()
                queue = _invoice_queue()

            self.assertEqual(sync_result["generated"], [source.name])
            self.assertGreater(sync_result["mapped_barcodes"], 0)
            queued_invoice = queue[0]
            for line in invoice.lines:
                if line.barcode:
                    self.assertEqual(queued_invoice["article_ids"][str(line.line_id)], 98765)

    def test_invoice_sync_passes_description_product_ids_to_workbook_generation(self) -> None:
        from facturacion.revisar_excepciones import _sync_invoice_workbooks

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            output_dir.mkdir()
            (xml_dir / source.name).write_bytes(source.read_bytes())
            description_ids = {"linea sin gtin": 54321}
            with (
                patch("facturacion.revisar_excepciones.DEFAULT_XML_DIR", xml_dir),
                patch("facturacion.revisar_excepciones.DEFAULT_OUTPUT_DIR", output_dir),
                patch("facturacion.revisar_excepciones.REVIEW_FILE", output_dir / "reviews.json"),
                patch("facturacion.revisar_excepciones._cached_barcode_product_ids", return_value={}),
                patch("facturacion.revisar_excepciones.get_product_ids_by_descriptions", return_value=description_ids) as lookup,
                patch("facturacion.revisar_excepciones._existing_tax_codes", return_value={}),
                patch("facturacion.revisar_excepciones.generate_workbook") as generate_workbook_mock,
            ):
                _sync_invoice_workbooks(force=True)

            lookup.assert_called_once()
            self.assertEqual(generate_workbook_mock.call_args.kwargs["product_ids_by_description"], description_ids)

    def test_sync_generates_invoice_offline_when_description_lookup_fails(self) -> None:
        from facturacion.revisar_excepciones import _sync_invoice_workbooks

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            output_dir.mkdir()
            (xml_dir / source.name).write_bytes(source.read_bytes())
            with (
                patch("facturacion.revisar_excepciones.DEFAULT_XML_DIR", xml_dir),
                patch("facturacion.revisar_excepciones.DEFAULT_OUTPUT_DIR", output_dir),
                patch("facturacion.revisar_excepciones.REVIEW_FILE", output_dir / "reviews.json"),
                patch("facturacion.revisar_excepciones._cached_barcode_product_ids", return_value={}),
                patch(
                    "facturacion.revisar_excepciones.get_product_ids_by_descriptions",
                    side_effect=RuntimeError("Azure firewall blocked the description lookup"),
                ),
                patch("facturacion.revisar_excepciones._existing_tax_codes", return_value={}),
            ):
                result = _sync_invoice_workbooks(force=True)

            self.assertIn("description lookup", result["lookup_error"])
            self.assertEqual(result["generated"], [source.name])
            self.assertEqual(result["mapped_barcodes"], 0)
            generated_files = list(output_dir.glob("*_effi.xlsx"))
            self.assertEqual(len(generated_files), 1)
            self.assertTrue(generated_files[0].is_file())

    def test_sync_preserves_existing_article_ids_when_description_lookup_fails(self) -> None:
        from facturacion.generar_importacion import generate_workbook
        from facturacion.revisar_excepciones import _invoice_output_path, _sync_invoice_workbooks

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        invoice = parse_invoice(source)
        line = next(line for line in invoice.lines if line.barcode)
        barcode = line.barcode
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            xml_dir = root / "xml"
            output_dir = root / "output"
            xml_dir.mkdir()
            output_dir.mkdir()
            output = output_dir / f"{invoice.invoice_id}_effi.xlsx"
            (xml_dir / source.name).write_bytes(source.read_bytes())
            initial_id = 98765
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={barcode: initial_id},
                product_ids_by_description={},
            )
            with (
                patch("facturacion.revisar_excepciones.DEFAULT_XML_DIR", xml_dir),
                patch("facturacion.revisar_excepciones.DEFAULT_OUTPUT_DIR", output_dir),
                patch("facturacion.revisar_excepciones.REVIEW_FILE", output_dir / "reviews.json"),
                patch("facturacion.revisar_excepciones._cached_barcode_product_ids", return_value={}),
                patch(
                    "facturacion.revisar_excepciones.get_product_ids_by_descriptions",
                    side_effect=RuntimeError("Azure firewall blocked the description lookup"),
                ),
                patch("facturacion.revisar_excepciones._existing_tax_codes", return_value={}),
            ):
                result = _sync_invoice_workbooks(force=True)
                workbook = load_workbook(output, read_only=True, data_only=True)
                try:
                    detail = workbook["Hoja1"]
                    self.assertEqual(detail.cell(2, 1).value, initial_id)
                finally:
                    workbook.close()

            self.assertIn("description lookup", result["lookup_error"])
            self.assertEqual(result["generated"], [])
            self.assertEqual(result["mapped_barcodes"], 0)

    def test_prorates_bonus_units_for_repeated_barcode(self) -> None:
        source = XML_DIR / "ad089030295502726356720260928112521401.xml"
        invoice = parse_invoice(source)
        repeated_lines = [line for line in invoice.lines if line.barcode == "7702113048636"]
        self.assertEqual([line.line_id for line in repeated_lines], ["2", "3"])
        total_quantity = sum((line.quantity for line in repeated_lines), Decimal("0"))
        total_net = sum((line.net_amount for line in repeated_lines), Decimal("0"))

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "bonus_by_barcode.xlsx"
            warnings = generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={"7702113048636": 98765},
                product_ids_by_description={},
            )
            workbook = load_workbook(output, data_only=False)
            detail = workbook["Hoja1"]
            exceptions = workbook["Excepciones"]
            headers = {exceptions.cell(1, column).value: column for column in range(1, exceptions.max_column + 1)}
            exception_rows = {
                str(exceptions.cell(row, headers["Línea"]).value): row
                for row in range(2, exceptions.max_row + 1)
            }
            import_rows = {
                int(exceptions.cell(exception_rows[line.line_id], headers["Fila importación"]).value)
                for line in repeated_lines
            }
            for line in repeated_lines:
                exception_row = exception_rows[line.line_id]
                self.assertEqual(
                    Decimal(str(exceptions.cell(exception_row, headers["Unidades totales del producto"]).value)),
                    total_quantity,
                )
                self.assertAlmostEqual(
                    float(exceptions.cell(exception_row, headers["Costo unitario efectivo"]).value),
                    float(total_net / total_quantity),
                    places=8,
                )
            self.assertEqual(len(import_rows), 1)
            import_row = import_rows.pop()
            self.assertEqual(detail.cell(import_row, 1).value, 98765)
            self.assertEqual(Decimal(str(detail.cell(import_row, 6).value)), total_quantity)
            self.assertEqual(Decimal(str(detail.cell(import_row, 8).value)), Decimal("0"))
            self.assertEqual(
                (Decimal(str(detail.cell(import_row, 7).value)) * total_quantity).quantize(Decimal("0.01")),
                total_net.quantize(Decimal("0.01")),
            )
            self.assertEqual(detail.cell(import_row, 9).value, 1)
            self.assertTrue(any("Bonificación consolidada" in warning for warning in warnings))
            workbook.close()

            from facturacion.revisar_excepciones import _reviewed_workbook

            reviewed = _reviewed_workbook(
                {"invoice": invoice, "source": source, "output": output},
                {
                    repeated_lines[0].line_id: {"status": "confirmed", "product_id": 98765},
                    repeated_lines[1].line_id: {"status": "no_match"},
                },
            )
            reviewed_workbook = load_workbook(reviewed, data_only=True)
            self.assertEqual(reviewed_workbook["Hoja1"].cell(import_row, 1).value, 98765)
            reviewed_workbook.close()

            cost_reviewed = _reviewed_workbook(
                {"invoice": invoice, "source": source, "output": output},
                {
                    repeated_lines[0].line_id: {
                        "status": "cost_saved",
                        "final_unit_cost": "8000",
                        "is_bonus": False,
                    },
                    repeated_lines[1].line_id: {
                        "status": "cost_saved",
                        "final_unit_cost": "8000",
                        "is_bonus": True,
                        "bonus_link_line": repeated_lines[0].line_id,
                    },
                },
            )
            cost_workbook = load_workbook(cost_reviewed, data_only=True)
            cost_detail = cost_workbook["Hoja1"]
            self.assertEqual(cost_detail.cell(import_row, 6).value, 27)
            self.assertEqual(cost_detail.cell(import_row, 7).value, 8000)
            self.assertEqual(cost_detail.cell(import_row, 8).value, 0)
            cost_workbook.close()

    def test_bonus_cost_summary_prorates_paid_net_across_linked_units(self) -> None:
        from facturacion.revisar_excepciones import _bonus_cost_summary

        invoice = parse_invoice(XML_DIR / "ad089030295502726356720260928112521401.xml")
        product_lines = [line for line in invoice.lines if line.barcode == "7702113048636"]
        bonus_line_ids = {str(line.line_id) for line in product_lines if line.net_amount == 0}
        total_quantity = sum((line.quantity for line in product_lines), Decimal("0"))
        total_net = sum((line.net_amount for line in product_lines), Decimal("0"))

        summary = _bonus_cost_summary(product_lines, bonus_line_ids)

        self.assertEqual(summary["total_quantity"], total_quantity)
        self.assertEqual(summary["net_total"], total_net)
        self.assertEqual(
            summary["billed_quantity"],
            sum((line.quantity for line in product_lines if line.net_amount > 0), Decimal("0")),
        )
        self.assertEqual(summary["bonus_quantity"], total_quantity - summary["billed_quantity"])
        self.assertEqual(summary["effective_unit_cost"], total_net / total_quantity)
        paid_line_id = next(str(line.line_id) for line in product_lines if line.net_amount > 0)
        with self.assertRaisesRegex(ValueError, "neto cero"):
            _bonus_cost_summary(product_lines, {paid_line_id})

    def test_prorates_bonus_units_without_barcode_using_item_reference(self) -> None:
        source = XML_DIR / "52089_e617b4bf4e04d5380af089e90fb2f16b17f43f96e25e3d78623dfd2b55c8ba9f70b5b06ff39c2fe285bc58524803dd5b.xml"
        invoice = parse_invoice(source)
        repeated_lines = [line for line in invoice.lines if line.item_code == "NOUSHAPROT"]
        self.assertEqual(len(repeated_lines), 2)
        self.assertTrue(all(not line.barcode for line in repeated_lines))
        total_quantity = sum((line.quantity for line in repeated_lines), Decimal("0"))
        total_net = sum((line.net_amount for line in repeated_lines), Decimal("0"))

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "bonus_by_reference.xlsx"
            generate_workbook(source, TEMPLATE, output, DEFAULT_TAX_CODES)
            workbook = load_workbook(output, data_only=False)
            detail = workbook["Hoja1"]
            exceptions = workbook["Excepciones"]
            headers = {exceptions.cell(1, column).value: column for column in range(1, exceptions.max_column + 1)}
            exception_rows = {
                str(exceptions.cell(row, headers["Línea"]).value): row
                for row in range(2, exceptions.max_row + 1)
            }
            import_rows = {
                int(exceptions.cell(exception_rows[line.line_id], headers["Fila importación"]).value)
                for line in repeated_lines
            }
            self.assertEqual(len(import_rows), 1)
            import_row = import_rows.pop()
            self.assertIsNone(detail.cell(import_row, 1).value)
            self.assertEqual(Decimal(str(detail.cell(import_row, 6).value)), total_quantity)
            self.assertEqual(Decimal(str(detail.cell(import_row, 7).value)), Decimal("13125"))
            self.assertEqual(Decimal(str(detail.cell(import_row, 8).value)), Decimal("0"))
            self.assertEqual(
                (Decimal(str(detail.cell(import_row, 7).value)) * total_quantity).quantize(Decimal("0.01")),
                total_net.quantize(Decimal("0.01")),
            )
            workbook.close()

    def test_reviewed_workbook_prorates_explicit_bonus_link_without_override(self) -> None:
        from facturacion.revisar_excepciones import _line_product_key, _reviewed_workbook

        source = XML_DIR / "ad089030295502726356720260928112521401.xml"
        invoice = parse_invoice(source)
        grouped_lines = [line for line in invoice.lines if line.barcode == "7702113048636"]
        paid_line = next(line for line in grouped_lines if line.net_amount > 0)
        bonus_line = next(line for line in grouped_lines if line.net_amount == 0)
        total_quantity = sum((line.quantity for line in grouped_lines), Decimal("0"))
        total_net = sum((line.net_amount for line in grouped_lines), Decimal("0"))

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "bonus_review.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={"7702113048636": 98765},
                product_ids_by_description={},
            )
            original_workbook = load_workbook(output, data_only=True)
            exception_sheet = original_workbook["Excepciones"]
            headers = {
                exception_sheet.cell(1, column).value: column
                for column in range(1, exception_sheet.max_column + 1)
            }
            import_row = next(
                int(exception_sheet.cell(exception_row, headers["Fila importación"]).value)
                for exception_row in range(2, exception_sheet.max_row + 1)
                if str(exception_sheet.cell(exception_row, headers["Línea"]).value) == bonus_line.line_id
            )
            original_workbook.close()
            reviewed = _reviewed_workbook(
                {"invoice": invoice, "source": source, "output": output},
                {
                    paid_line.line_id: {"status": "confirmed", "product_id": 98765},
                    bonus_line.line_id: {
                        "status": "confirmed",
                        "product_id": 98765,
                        "is_bonus": True,
                        "bonus_link_line": paid_line.line_id,
                    },
                },
            )
            workbook = load_workbook(reviewed, data_only=True)
            detail = workbook["Hoja1"]
            self.assertEqual(detail.cell(import_row, 1).value, 98765)
            self.assertEqual(Decimal(str(detail.cell(import_row, 6).value)), total_quantity)
            self.assertAlmostEqual(
                float(detail.cell(import_row, 7).value),
                float(total_net / total_quantity),
                places=8,
            )
            self.assertEqual(Decimal(str(detail.cell(import_row, 8).value)), Decimal("0"))
            workbook.close()

            incompatible_paid_line = next(
                line for line in invoice.lines
                if line.net_amount > 0 and _line_product_key(line) != _line_product_key(bonus_line)
            )
            with self.assertRaisesRegex(ValueError, "producto incompatible"):
                _reviewed_workbook(
                    {"invoice": invoice, "source": source, "output": output},
                    {
                        incompatible_paid_line.line_id: {"status": "confirmed", "product_id": 65432},
                        bonus_line.line_id: {
                            "status": "confirmed",
                            "product_id": 98765,
                            "is_bonus": True,
                            "bonus_link_line": incompatible_paid_line.line_id,
                        },
                    },
                )

    def test_final_imported_net_matches_all_sample_invoice_totals(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            for source in sorted(XML_DIR.glob("*.xml")):
                invoice = parse_invoice(source)
                output = Path(temp_dir) / f"{source.stem}.xlsx"
                generate_workbook(
                    source,
                    TEMPLATE,
                    output,
                    DEFAULT_TAX_CODES,
                    product_ids_by_barcode={},
                    product_ids_by_description={},
                )
                workbook = load_workbook(output, data_only=True)
                detail = workbook.active
                imported_total = sum(
                    (
                        Decimal(str(detail.cell(row, 6).value or 0))
                        * Decimal(str(detail.cell(row, 7).value or 0))
                        - Decimal(str(detail.cell(row, 8).value or 0))
                        for row in range(2, detail.max_row + 1)
                    ),
                    Decimal("0"),
                )
                invoice_total = sum((line.net_amount for line in invoice.lines), Decimal("0"))
                self.assertEqual(
                    imported_total.quantize(Decimal("0.01")),
                    invoice_total.quantize(Decimal("0.01")),
                    source.name,
                )
                workbook.close()

    def test_exception_cursor_advances_to_next_pending_line(self) -> None:
        from facturacion.revisar_excepciones import _next_pending_index

        exceptions = [{"Línea": str(index)} for index in range(1, 5)]
        reviews = {
            "1": {"status": "confirmed"},
            "2": {"status": "no_match"},
        }
        self.assertEqual(_next_pending_index(exceptions, reviews, 0), 2)
        self.assertEqual(_next_pending_index(exceptions, reviews, 3), 3)
        reviews.update({"3": {"status": "confirmed"}, "4": {"status": "no_match"}})
        self.assertIsNone(_next_pending_index(exceptions, reviews, 0))

    def test_invoice_cursor_advances_to_next_invoice_with_pending_lines(self) -> None:
        from facturacion.revisar_excepciones import _next_pending_invoice_index

        queue = [
            {"source": Path("first.xml"), "exceptions": [{"Línea": "1"}]},
            {"source": Path("second.xml"), "exceptions": [{"Línea": "1"}]},
        ]
        state = {"first.xml": {"1": {"status": "confirmed"}}}
        self.assertEqual(_next_pending_invoice_index(queue, state, 1), 1)
        state["second.xml"] = {"1": {"status": "no_match"}}
        self.assertIsNone(_next_pending_invoice_index(queue, state, 1))

    def test_invoice_approval_requires_a_decision_for_every_exception(self) -> None:
        from facturacion.revisar_excepciones import _invoice_review_complete

        queue_item = {"exceptions": [{"Línea": "1"}, {"Línea": "2"}]}
        self.assertFalse(_invoice_review_complete(queue_item, {"1": {"status": "confirmed"}}))
        self.assertFalse(_invoice_review_complete(queue_item, {"1": {"status": "cost_saved"}, "2": {"status": "no_match"}}))
        self.assertTrue(_invoice_review_complete(queue_item, {
            "1": {"status": "confirmed"},
            "2": {"status": "no_match"},
        }))

    def test_automatic_bonus_does_not_block_invoice_approval_or_cursor(self) -> None:
        from facturacion.revisar_excepciones import (
            _invoice_review_complete,
            _next_pending_index,
            _product_review_status,
        )

        automatic_bonus = {"Línea": "1", "Tipo": "Bonificación/prorrateo: costo efectivo calculado"}
        unresolved_gtin = {"Línea": "2", "Tipo": "Código de barras no registrado en EFFI"}
        exceptions = [automatic_bonus, unresolved_gtin]
        queue_item = {"exceptions": exceptions}

        self.assertFalse(_invoice_review_complete(queue_item, {}))
        self.assertEqual(_next_pending_index(exceptions, {}, 0), 1)
        self.assertTrue(_invoice_review_complete(queue_item, {"2": {"status": "no_match"}}))
        self.assertTrue(_invoice_review_complete({"exceptions": [automatic_bonus]}, {}))
        self.assertIsNone(_next_pending_index([automatic_bonus], {}, 0))
        self.assertEqual(
            _product_review_status({"line_ids": ["1"]}, {"1": automatic_bonus}, {}),
            "Validado automáticamente",
        )

    def test_invoice_approval_expires_when_source_or_decisions_change(self) -> None:
        from facturacion.revisar_excepciones import (
            INVOICE_APPROVAL_KEY,
            _approval_fingerprint,
            _invoice_approval_is_current,
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "invoice.xml"
            source.write_text("<Invoice />", encoding="utf-8")
            reviews = {"1": {"status": "confirmed", "product_id": 123}}
            reviews[INVOICE_APPROVAL_KEY] = {
                "status": "approved",
                "fingerprint": _approval_fingerprint(source, reviews),
            }
            self.assertTrue(_invoice_approval_is_current(source, reviews))

            reviews["1"]["product_id"] = 456
            self.assertFalse(_invoice_approval_is_current(source, reviews))
            reviews["1"]["product_id"] = 123
            source.write_text("<Invoice id='changed' />", encoding="utf-8")
            self.assertFalse(_invoice_approval_is_current(source, reviews))

    def test_invalidated_reviewed_workbook_is_versioned(self) -> None:
        from facturacion.revisar_excepciones import _invalidate_reviewed_workbook

        with tempfile.TemporaryDirectory() as temp_dir:
            reviewed = Path(temp_dir) / "invoice_revisado.xlsx"
            reviewed.write_bytes(b"old workbook")

            backup = _invalidate_reviewed_workbook(reviewed)

            self.assertIsNotNone(backup)
            self.assertFalse(reviewed.exists())
            self.assertTrue(backup.exists())
            self.assertTrue(backup.name.startswith("invoice_revisado_invalidado_"))

    def test_decision_change_invalidates_approval_and_archives_reviewed_workbook(self) -> None:
        from facturacion import revisar_excepciones
        from facturacion.revisar_excepciones import INVOICE_APPROVAL_KEY, _record_review

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_output = Path(temp_dir)
            source = temp_output / "invoice.xml"
            source.write_text("<Invoice />", encoding="utf-8")
            state = {source.name: {INVOICE_APPROVAL_KEY: {"status": "approved"}, "1": {"status": "confirmed"}}}
            output = temp_output / "invoice_effi.xlsx"
            output.write_bytes(b"base workbook")
            reviewed = output.with_name(output.stem.removesuffix("_effi") + "_revisado.xlsx")
            reviewed.write_bytes(b"reviewed workbook")

            with patch.object(revisar_excepciones, "DEFAULT_OUTPUT_DIR", temp_output):
                _record_review(state, source, "1", "confirmed", None, "Ana", invoice_id="invoice")

            self.assertNotIn(INVOICE_APPROVAL_KEY, state[source.name])
            self.assertFalse(reviewed.exists())
            self.assertTrue(any(
                backup.name.startswith("invoice_revisado_invalidado_")
                for backup in temp_output.glob("*_revisado_invalidado_*.xlsx")
            ))

    def test_dashboard_metrics_and_product_queue_use_invoice_review_data(self) -> None:
        from facturacion.revisar_excepciones import (
            _dashboard_metrics,
            _invoice_queue,
            _pending_review_count,
            _product_review_rows,
        )

        queue = _invoice_queue()
        metrics = _dashboard_metrics(queue, {})
        product_rows = _product_review_rows(queue, {})
        expected_pending = sum(_pending_review_count(item, {}) for item in queue)
        expected_gtin_exceptions = sum(
            bool(line.barcode) and not item["article_ids"].get(str(line.line_id))
            for item in queue
            for line in item["invoice"].lines
        )

        self.assertEqual(metrics["pending_products"], expected_pending)
        self.assertEqual(metrics["pending_invoices"], sum(_pending_review_count(item, {}) > 0 for item in queue))
        self.assertEqual(metrics["gtin_exceptions"], expected_gtin_exceptions)
        self.assertEqual(len(product_rows), sum(len(item["exceptions"]) for item in queue))
        self.assertTrue(all(row["Estado"] in {"Pendiente", "Automático"} for row in product_rows))

    def test_review_history_includes_line_decisions_and_invoice_approval(self) -> None:
        from facturacion.revisar_excepciones import (
            INVOICE_APPROVAL_KEY,
            INVOICE_APPROVAL_HISTORY_KEY,
            _approval_fingerprint,
            _history_rows,
            _invoice_queue,
        )

        queue = _invoice_queue()
        source_name = queue[0]["source"].name
        invoice_reviews = {
            "1": {
                "status": "confirmed",
                "product_name": "Producto confirmado",
                "reviewed_at": "2026-10-02T09:00:00+00:00",
            },
            INVOICE_APPROVAL_KEY: {
                "status": "approved",
                "approved_by": "Camila",
                "approved_at": "2026-10-02T09:30:00+00:00",
            },
        }
        invoice_reviews[INVOICE_APPROVAL_HISTORY_KEY] = [
            {
                "status": "approved",
                "approved_by": "Ana",
                "approved_at": "2026-10-02T08:30:00+00:00",
                "fingerprint": "previous-version",
            },
            invoice_reviews[INVOICE_APPROVAL_KEY],
        ]
        invoice_reviews[INVOICE_APPROVAL_KEY]["fingerprint"] = _approval_fingerprint(
            queue[0]["source"], invoice_reviews
        )
        rows = _history_rows(queue, {source_name: invoice_reviews})

        self.assertEqual(len(rows), 3)
        self.assertEqual({row["Estado"] for row in rows}, {"Confirmado", "Aprobada", "Aprobación anterior"})
        self.assertEqual({row["Responsable"] for row in rows}, {"Ana", "Camila", ""})

    def test_no_match_review_clears_article_in_reviewed_workbook(self) -> None:
        from facturacion.revisar_excepciones import _reviewed_workbook

        source = XML_DIR / "ad089030295502726356720260928112521401.xml"
        invoice = parse_invoice(source)
        line = invoice.lines[0]
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "invoice_effi.xlsx"
            generate_workbook(source, TEMPLATE, output, DEFAULT_TAX_CODES)
            reviewed = _reviewed_workbook(
                {"invoice": invoice, "source": source, "output": output},
                {line.line_id: {"status": "no_match"}},
            )
            workbook = load_workbook(reviewed, data_only=True)
            self.assertIsNone(workbook.active.cell(2, 1).value)
            workbook.close()

    def test_writes_only_importable_detail_fields_and_preserves_header(self) -> None:
        source = XML_DIR / "ad090136037002826FEMJ00003856.xml"
        invoice = parse_invoice(source)
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "result.xlsx"
            generate_workbook(source, TEMPLATE, output, DEFAULT_TAX_CODES)
            workbook = load_workbook(output, data_only=False)
            sheet = workbook.active
            self.assertEqual(sheet.cell(1, 1).value, "Artículo (ID EFFI | Código de barras GTIN | Serie)")
            self.assertIsNone(sheet.cell(2, 1).value)
            self.assertEqual(sheet.cell(2, 6).value, 2)
            self.assertEqual(sheet.cell(2, 7).value, 29148.74)
            self.assertEqual(sheet.cell(2, 8).value, 11659.5)
            self.assertEqual(sheet.cell(2, 9).value, 1)
            self.assertEqual(sheet.cell(2, 2).value, 1)
            self.assertEqual(
                round(sheet.cell(2, 7).value * sheet.cell(2, 6).value - sheet.cell(2, 8).value, 2),
                round(float(invoice.lines[0].net_amount), 2),
            )
            workbook.close()

    def test_maps_barcode_to_effi_product_id(self) -> None:
        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        invoice = parse_invoice(source)
        barcode = invoice.lines[0].barcode
        self.assertTrue(barcode)

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "mapped.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={barcode: 98765},
            )
            workbook = load_workbook(output, data_only=False)
            self.assertEqual(workbook.active.cell(2, 1).value, 98765)
            self.assertIsNone(workbook.active.cell(2, 2).value)
            workbook.close()

    def test_unregistered_barcode_does_not_fall_back_to_description(self) -> None:
        from facturacion.generar_importacion import _normalize_description

        source = XML_DIR / "ad080021284002726837920260926084002509.xml"
        invoice = parse_invoice(source)
        line = next(item for item in invoice.lines if item.barcode)
        candidate = {
            "IdProducto": 99999,
            "Nombre": line.description,
            "Referencia": "TEST",
            "Marca": "TEST",
            "Categoria": "TEST",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "barcode_not_registered.xlsx"
            warnings = generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={},
                product_ids_by_description={_normalize_description(line.description): [candidate]},
            )
            workbook = load_workbook(output, data_only=False)
            self.assertIsNone(workbook.active.cell(2, 1).value)
            exception = workbook["Excepciones"]
            self.assertIn("Código de barras no registrado en EFFI", exception.cell(2, 2).value)
            self.assertIn(line.description, exception.cell(2, 2).value)
            barcode_warning = next(warning for warning in warnings if "Código de barras" in warning)
            self.assertIn(line.description, barcode_warning)
            self.assertIn("vincular este código en EFFI", barcode_warning)
            workbook.close()

    def test_maps_description_to_effi_product_id_when_barcode_missing(self) -> None:
        source = XML_DIR / "ad09013821010002600127292.xml"
        invoice = parse_invoice(source)
        self.assertTrue(invoice.lines)
        self.assertTrue(all(not line.barcode for line in invoice.lines))
        line = invoice.lines[0]
        candidate = {
            "IdProducto": 5001,
            "Nombre": line.description,
            "Referencia": "HK IGORA",
            "Marca": "IGORA",
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "mapped_by_description.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={},
                product_ids_by_description={"hk igora vit n1 1 0 negro px24": [candidate]},
            )
            workbook = load_workbook(output, data_only=False)
            self.assertEqual(workbook.active.cell(2, 1).value, 5001)
            workbook.close()

    def test_description_match_prefers_brand_presentation_and_price(self) -> None:
        from facturacion.generar_importacion import _candidate_score

        line = parse_invoice(XML_DIR / "ad09013821010002600127292.xml").lines[0]
        matching = {
            "IdProducto": 5001,
            "Nombre": "IGORA VITAL N1 1-0 NEGRO 60 ML",
            "Referencia": "HK IGORA N1 1-0",
            "Marca": "IGORA",
            "Categoria": "TINTURAS",
            "PrecioConImpuestoTarifaNormal": 14920,
            "PrecioConImpuestoTarifaMayorista": 13000,
        }
        wrong_product = {
            **matching,
            "IdProducto": 5002,
            "Nombre": "IGORA VITAL N3 4-0 CASTAÑO 60 ML",
            "Referencia": "IGORA N3 4-0",
            "Marca": "OTRA MARCA",
            "PrecioConImpuestoTarifaNormal": 30000,
        }
        self.assertGreater(_candidate_score(line, matching), _candidate_score(line, wrong_product))

    def test_description_match_rejects_equal_scoring_candidates(self) -> None:
        from facturacion.generar_importacion import _resolved_product

        line = parse_invoice(XML_DIR / "ad09013821010002600127292.xml").lines[0]
        candidate = {
            "IdProducto": 5001,
            "Nombre": line.description,
            "Referencia": "HK IGORA",
            "Marca": "IGORA",
        }
        duplicate = {**candidate, "IdProducto": 5002}

        resolved_id, best, score, margin, accepted = _resolved_product(
            line,
            None,
            {"hk igora vit n1 1 0 negro px24": [candidate, duplicate]},
        )

        self.assertIsNone(resolved_id)
        self.assertEqual(best, candidate)
        self.assertGreater(score, 0.72)
        self.assertEqual(margin, 0.0)
        self.assertFalse(accepted)

    def test_auto_match_without_invoice_barcode_writes_catalog_barcode(self) -> None:
        from facturacion.generar_importacion import _normalize_description

        source = next(
            path for path in XML_DIR.glob("*.xml")
            if parse_invoice(path).invoice_id == "FECG315045"
        )
        invoice = parse_invoice(source)
        line = next(line for line in invoice.lines if "12.013" in line.description)
        candidate = {
            "IdProducto": 5546,
            "CodigoBarras": "7701234567890",
            "Nombre": "TNT. NOUVELLE #12.013",
            "Referencia": "005563",
            "Marca": "EURO STYLE",
            "Categoria": "TINTURAS",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "automatic_match.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={},
                product_ids_by_description={_normalize_description(line.description): [candidate]},
            )
            workbook = load_workbook(output, data_only=True)
            matches = workbook["Coincidencias automáticas"]
            row = next(
                row for row in matches.iter_rows(min_row=2, values_only=True)
                if str(row[0]) == line.line_id
            )
            self.assertEqual(row[2], 5546)
            self.assertEqual(row[4], "7701234567890")
            self.assertIn(5546, [workbook.active.cell(index, 1).value for index in range(2, workbook.active.max_row + 1)])
            workbook.close()

    def test_amount_discrepancy_and_closest_candidate_are_written_to_exceptions(self) -> None:
        from facturacion.generar_importacion import _normalize_description

        source = XML_DIR / "ad09013821010002600127292.xml"
        invoice = parse_invoice(source)
        line = next(item for item in invoice.lines if item.line_id == "7")
        closest = {
            "IdProducto": 777,
            "Nombre": "PRODUCTO CERCANO PARA REVISION",
            "Referencia": "REF-777",
            "Marca": "MARCA",
            "Categoria": "CATEGORIA",
            "PrecioConImpuestoTarifaNormal": 14000,
            "PrecioConImpuestoTarifaMayorista": 12000,
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "reconciled.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={},
                product_ids_by_description={_normalize_description(line.description): [closest]},
            )
            workbook = load_workbook(output, data_only=False)
            detail = workbook["Hoja1"]
            row_number = int(line.line_id) + 1
            self.assertEqual(
                round(
                    detail.cell(row_number, 7).value * detail.cell(row_number, 6).value
                    - detail.cell(row_number, 8).value,
                    2,
                ),
                round(float(line.net_amount), 2),
            )
            exceptions = workbook["Excepciones"]
            exception_row = next(
                row for row in range(2, exceptions.max_row + 1)
                if exceptions.cell(row, 1).value == line.line_id
            )
            self.assertEqual(exceptions.cell(exception_row, 12).value, 777)
            self.assertNotEqual(exceptions.cell(exception_row, 11).value, 0)
            workbook.close()

    def test_all_sample_xml_files_have_invoice_lines(self) -> None:
        invoices = [parse_invoice(path) for path in XML_DIR.glob("*.xml")]
        self.assertEqual(len(invoices), 8)
        self.assertEqual(sum(len(invoice.lines) for invoice in invoices), 170)


if __name__ == "__main__":
    unittest.main()