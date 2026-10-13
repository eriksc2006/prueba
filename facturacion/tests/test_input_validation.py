from __future__ import annotations

import argparse
import csv
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from facturacion.generar_importacion import (
    DEFAULT_TAX_CODES,
    InvoiceParseError,
    _parse_tax_mapping,
    parse_invoice,
)


class InvoiceInputValidationTests(unittest.TestCase):
    def test_rejects_malformed_xml(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "malformed.xml"
            source.write_text("<Invoice>", encoding="utf-8")

            with self.assertRaisesRegex(InvoiceParseError, "No se pudo leer el XML"):
                parse_invoice(source)

    def test_rejects_invoice_without_detail_lines(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "empty.xml"
            source.write_text("<Invoice />", encoding="utf-8")

            with self.assertRaisesRegex(InvoiceParseError, "no contiene líneas"):
                parse_invoice(source)

    def test_rejects_xml_larger_than_configured_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "large.xml"
            source.write_text("<Invoice />", encoding="utf-8")

            with patch("facturacion.generar_importacion.MAX_XML_BYTES", 4):
                with self.assertRaisesRegex(InvoiceParseError, "supera el límite"):
                    parse_invoice(source)

    def test_rejects_xml_exceeding_configured_depth(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "deep.xml"
            source.write_text("<Invoice><a><b><c /></b></a></Invoice>", encoding="utf-8")

            with patch("facturacion.generar_importacion.MAX_XML_DEPTH", 3):
                with self.assertRaisesRegex(InvoiceParseError, "profundidad máxima"):
                    parse_invoice(source)

    def test_rejects_xml_exceeding_configured_element_count(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "many-elements.xml"
            source.write_text("<Invoice><a /><b /></Invoice>", encoding="utf-8")

            with patch("facturacion.generar_importacion.MAX_XML_ELEMENTS", 2):
                with self.assertRaisesRegex(InvoiceParseError, "número máximo de elementos"):
                    parse_invoice(source)

    def test_rejects_non_finite_invoice_amounts(self) -> None:
        xml = (
            "<Invoice><InvoiceLine><ID>1</ID>"
            "<InvoicedQuantity>NaN</InvoicedQuantity>"
            "<LineExtensionAmount>0</LineExtensionAmount>"
            "</InvoiceLine></Invoice>"
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "non_finite.xml"
            source.write_text(xml, encoding="utf-8")

            with self.assertRaisesRegex(InvoiceParseError, "Valor no finito"):
                parse_invoice(source)


class TaxMappingTests(unittest.TestCase):
    def test_adds_and_overrides_tax_codes(self) -> None:
        mapping = _parse_tax_mapping(["5=3", "19=9"])

        self.assertEqual(mapping, {Decimal("19"): 9, Decimal("5"): 3})
        self.assertEqual(DEFAULT_TAX_CODES, {Decimal("19"): 1})

    def test_rejects_invalid_tax_mapping(self) -> None:
        for value in ("19", "abc=1", "19=not-an-int"):
            with self.subTest(value=value):
                with self.assertRaises(argparse.ArgumentTypeError):
                    _parse_tax_mapping([value])


class CsvExportSecurityTests(unittest.TestCase):
    def test_csv_export_neutralizes_formula_like_catalog_values(self) -> None:
        from facturacion import azure_codes

        codes = [
            "7701234567890",
            "=1+1",
            "  =1+1",
            "\t=1+1",
            "+1+1",
            "-1+1",
            "@SUM(A1:A2)",
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "codes.csv"
            with (
                patch("facturacion.azure_codes.get_codes", return_value=codes),
                patch("sys.argv", ["azure_codes", "--output", str(output)]),
            ):
                self.assertEqual(azure_codes.main(), 0)

            with output.open(newline="", encoding="utf-8-sig") as source:
                rows = list(csv.reader(source))

        self.assertEqual(rows[0], ["CodigoBarras"])
        self.assertEqual(
            [row[0] for row in rows[1:]],
            ["7701234567890", "'=1+1", "'  =1+1", "'\t=1+1", "'+1+1", "'-1+1", "'@SUM(A1:A2)"],
        )


if __name__ == "__main__":
    unittest.main()