from __future__ import annotations

import tempfile
import unittest
from decimal import Decimal
from pathlib import Path

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

    def test_rejects_xml_without_invoice(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "bad.xml"
            path.write_text("<Root />", encoding="utf-8")
            with self.assertRaises(InvoiceParseError):
                parse_invoice(path)


class WorkbookGenerationTests(unittest.TestCase):
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
            self.assertEqual(sheet.cell(2, 1).value, "MJPT-235")
            self.assertEqual(sheet.cell(2, 6).value, 2)
            self.assertEqual(sheet.cell(2, 7).value, 29148.74)
            self.assertEqual(sheet.cell(2, 8).value, 11659.5)
            self.assertEqual(sheet.cell(2, 9).value, 1)
            self.assertIsNone(sheet.cell(2, 2).value)
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

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "mapped_by_description.xlsx"
            generate_workbook(
                source,
                TEMPLATE,
                output,
                DEFAULT_TAX_CODES,
                product_ids_by_barcode={},
                product_ids_by_description={"hk igora vit n1 1 0 negro px24": 5001},
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
        self.assertEqual(len(invoices), 7)
        self.assertEqual(sum(len(invoice.lines) for invoice in invoices), 142)


if __name__ == "__main__":
    unittest.main()