"""Regression tests for the business rules owned by the Python matcher."""

from __future__ import annotations

import sys
import csv
import json
import os
import subprocess
import tempfile
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
from engine import GrRow, Invoice, Settings, VatLine, match_batch  # noqa: E402


def gr(source: str, receipt: str, material: str, qty: str, day: int) -> GrRow:
    return GrRow(source, receipt, material, "LTV", Decimal(qty), date(2026, 9, day), "")


def invoice(number: str = "00000001") -> Invoice:
    return Invoice("INV-" + number, "PDF-0001", number, date(2026, 9, 10), "LTV", "2500645835", 1, 1)


def line(inv: Invoice, seq: int, material: str, qty: str) -> VatLine:
    return VatLine(f"VL-{seq}", inv.invoice_id, seq, material, material, qty, Decimal(qty))


class MatcherRulesTests(unittest.TestCase):
    def test_fixed_runtime_entry_point_handles_an_empty_pdf_folder(self) -> None:
        with tempfile.TemporaryDirectory(prefix="vat-matcher-test-") as temp_root:
            root = Path(temp_root)
            input_dir, output_dir, pdf_dir = root / "input", root / "output", root / "pdf"
            input_dir.mkdir()
            pdf_dir.mkdir()
            (input_dir / "run_config.json").write_text(json.dumps({"run_id": "T-CLI", "pdf_folder": str(pdf_dir)}), encoding="utf-8")
            table_headers = {
                "config.csv": ["Key", "Value", "Description"],
                "gr_data.csv": ["SourceRow", "ReceiptNo", "Material", "Vendor", "QtyDoc", "QtyActual", "QtyMatch", "IB", "GRDate", "Flags"],
                "ncc_map.csv": ["CanonicalVendor", "TaxCode", "RawAlias", "FileAlias", "ParserProfile", "Active"],
                "parser_profiles.csv": ["ProfileID", "CanonicalVendor", "TaxCode", "InvoiceNoPattern", "DatePattern", "LinePattern", "Status", "SamplePdf", "Notes"],
                "material_scope_map.csv": ["MaterialNorm", "ScopeStatus", "Note"],
                "material_ncc_map.csv": ["MaterialNorm", "CanonicalVendor", "Source", "Active", "Note"],
            }
            for filename, headers in table_headers.items():
                with (input_dir / filename).open("w", encoding="utf-8-sig", newline="") as handle:
                    csv.writer(handle).writerow(headers)
            environment = os.environ.copy()
            environment["VAT_MATCHER_PYTHON"] = sys.executable
            launcher = Path(__file__).resolve().parents[1] / "python" / "run_tool.bat"
            completed = subprocess.run(
                ["cmd.exe", "/d", "/c", str(launcher), "--input", str(input_dir), "--output", str(output_dir)],
                env=environment, capture_output=True, text=True, check=False,
            )
            self.assertEqual(0, completed.returncode, completed.stderr)
            self.assertTrue((output_dir / "engine_result.json").exists())
            self.assertTrue((output_dir / "invoice_report.csv").exists())

    def test_one_vat_code_can_sum_three_receipts(self) -> None:
        inv = invoice()
        result = match_batch(
            [inv], [line(inv, 1, "BIW100", "1000")],
            [gr("1", "5001", "BIW100", "300", 5), gr("2", "5002", "BIW100", "200", 10), gr("3", "5003", "BIW100", "500", 12)],
            set(), Settings(), "T-1",
        )
        report = result["invoice_report"][0]
        self.assertEqual("MATCHED", report["Status"])
        self.assertEqual("5001+5002+5003", report["ProposedReceipts"])
        self.assertEqual(3, len(result["allocations"]))
        self.assertIn("Khớp 1/1 mã", report["Note"])

    def test_asymmetric_date_rule_excludes_six_days_before_and_three_days_after(self) -> None:
        inv = invoice()
        result = match_batch(
            [inv], [line(inv, 1, "BIW200", "100")],
            [gr("1", "4990", "BIW200", "100", 4), gr("2", "4991", "BIW200", "100", 13)],
            set(), Settings(gr_before_invoice_days=5, gr_after_invoice_days=2), "T-2",
        )
        report = result["invoice_report"][0]
        self.assertEqual("SUSPECT", report["Status"])
        self.assertIn("còn 1 mã BIW200 không tìm thấy phiếu phù hợp", report["Note"])

    def test_prefers_receipt_covering_more_codes_before_date_tie_breaker(self) -> None:
        inv = invoice()
        result = match_batch(
            [inv], [line(inv, 1, "BIWA", "100"), line(inv, 2, "BIWB", "50")],
            [
                gr("1", "5100", "BIWA", "100", 9), gr("2", "5100", "BIWB", "50", 9),
                gr("3", "5200", "BIWA", "100", 10), gr("4", "5300", "BIWB", "50", 10),
            ],
            set(), Settings(), "T-3",
        )
        self.assertEqual("5100", result["invoice_report"][0]["ProposedReceipts"])

    def test_a_gr_source_row_cannot_be_reused_by_a_second_invoice(self) -> None:
        first, second = invoice("00000001"), invoice("00000002")
        result = match_batch(
            [first, second], [line(first, 1, "BIW300", "100"), line(second, 1, "BIW300", "100")],
            [gr("1", "6001", "BIW300", "100", 10)], set(), Settings(), "T-4",
        )
        statuses = [row["Status"] for row in result["invoice_report"]]
        self.assertEqual(["MATCHED", "SUSPECT"], statuses)
        self.assertEqual(1, len(result["allocations"]))


if __name__ == "__main__":
    unittest.main()
