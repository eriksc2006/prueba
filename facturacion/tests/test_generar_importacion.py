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
    def test_writes_only_importable_detail_fields_and_preserves_header(self) -> None:
        source = XML_DIR / "ad090136037002826FEMJ00003856.xml"
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
            workbook.close()

    def test_all_sample_xml_files_have_invoice_lines(self) -> None:
        invoices = [parse_invoice(path) for path in XML_DIR.glob("*.xml")]
        self.assertEqual(len(invoices), 6)
        self.assertEqual(sum(len(invoice.lines) for invoice in invoices), 101)


if __name__ == "__main__":
    unittest.main()