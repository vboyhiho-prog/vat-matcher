#!/usr/bin/env python3
"""VAT Matcher Python engine.

The workbook is deliberately only the user interface: it exports normalised
tables to CSV, this module parses PDFs and matches GR rows, then writes the
same normalised tables back to CSV for Excel to display.  It never renames a
PDF and it never edits the source tracking workbook.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Iterable

try:
    import pymupdf
except ImportError as exc:  # surfaced by run_tool.bat before Excel imports anything
    raise SystemExit(
        "PyMuPDF is required. Install the approved fixed environment, then set "
        "VAT_MATCHER_PYTHON to that Python executable."
    ) from exc


EPSILON = Decimal("0.0001")
MAX_PLAN_ROWS = 220
MAX_DP_STATES = 30000
MAX_PLANS_PER_TOTAL = 8


PDF_FIELDS = ["PdfID", "OriginalPath", "Fingerprint", "ExtractStatus", "InvoiceCount"]
INVOICE_FIELDS = [
    "InvoiceID", "PdfID", "PageFrom", "PageTo", "InvoiceNo", "InvoiceDate",
    "VendorCanonical", "TaxCode", "ParseStatus",
]
VAT_LINE_FIELDS = [
    "VatLineID", "InvoiceID", "Seq", "MaterialRaw", "MaterialNorm", "QtyRaw",
    "Qty", "Unit", "ParseConfidence", "ScopeStatus",
]
CANDIDATE_FIELDS = [
    "CandidateID", "InvoiceID", "ReceiptSet", "VendorScore", "MaterialScore",
    "QtyScore", "DateScore", "HintScore", "TotalScore", "Rank", "Reasons",
]
ALLOCATION_FIELDS = [
    "AllocationID", "InvoiceID", "VatLineID", "ReceiptNo", "SourceRow",
    "AllocatedQty", "Residual", "Status",
]
REPORT_FIELDS = [
    "RunID", "PDF", "InvoiceNo", "Vendor", "InvoiceDate", "ProposedReceipts",
    "TotalScore", "Status", "RenamePreview", "Decision", "Note",
]
RECEIPT_REPORT_FIELDS = ["ReceiptNo", "Vendor", "GRDate", "LinkedInvoices", "Coverage", "Status"]
LOG_FIELDS = [
    "Timestamp", "RunID", "Severity", "Module", "Procedure", "Code", "Message",
    "OriginalPath", "NewPath",
]


@dataclass(frozen=True)
class Settings:
    """Business-rule settings.  `gr_before_invoice_days` is GR before VAT."""

    gr_before_invoice_days: int = 5
    gr_after_invoice_days: int = 2

    @classmethod
    def from_config(cls, config: dict[str, Any]) -> "Settings":
        def positive_int(name: str, default: int) -> int:
            try:
                value = int(Decimal(str(config.get(name, default))))
            except (InvalidOperation, ValueError):
                return default
            return value if value >= 0 else default

        return cls(
            gr_before_invoice_days=positive_int("GRBeforeInvoiceDays", 5),
            gr_after_invoice_days=positive_int("GRAfterInvoiceDays", 2),
        )


@dataclass
class GrRow:
    source_row: str
    receipt_no: str
    material: str
    vendor: str
    qty: Decimal
    gr_date: date | None
    flags: str


@dataclass
class Invoice:
    invoice_id: str
    pdf_id: str
    invoice_no: str
    invoice_date: date | None
    vendor: str
    tax_code: str
    page_from: int
    page_to: int
    parse_status: str = "PARSED"


@dataclass
class VatLine:
    vat_line_id: str
    invoice_id: str
    seq: int
    material_raw: str
    material: str
    qty_raw: str
    qty: Decimal
    unit: str = "EA"
    confidence: str = "0.9"
    scope_status: str = "UNKNOWN_MATERIAL"


@dataclass(frozen=True)
class SourceChoice:
    source_row: str
    receipt_no: str
    qty: Decimal
    distance: int


@dataclass
class ExactPlan:
    choices: tuple[SourceChoice, ...]

    @property
    def receipts(self) -> frozenset[str]:
        return frozenset(choice.receipt_no for choice in self.choices)

    @property
    def date_distance(self) -> int:
        return sum(choice.distance for choice in self.choices)


def utc_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalise(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def as_text(value: Any) -> str:
    return str(value or "").strip()


def decimal_value(value: Any) -> Decimal:
    raw = as_text(value).replace(" ", "")
    if not raw:
        return Decimal(0)
    # CSV may contain either Excel numeric values (45901 / 1000.5) or a
    # locale-formatted text value (1.000,5).  Preserve a decimal point when
    # both punctuation marks are present.
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(",", ".")
    try:
        return Decimal(raw)
    except InvalidOperation:
        return Decimal(0)


def quantity_text(value: Decimal) -> str:
    text = format(value.normalize(), "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def parse_date(value: Any) -> date | None:
    raw = as_text(value)
    if not raw:
        return None
    try:
        number = Decimal(raw.replace(",", "."))
        # Excel's epoch, accounting for the intentional 1900 leap-year bug.
        if Decimal(20000) < number < Decimal(100000):
            return date(1899, 12, 30) + timedelta(days=int(number))
    except InvalidOperation:
        pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(raw[:10], fmt).date()
        except ValueError:
            continue
    return None


def date_text(value: date | None) -> str:
    return value.isoformat() if value else ""


def read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, fieldnames: list[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({key: serialise(row.get(key, "")) for key in fieldnames})


def serialise(value: Any) -> str:
    if isinstance(value, Decimal):
        return quantity_text(value)
    if isinstance(value, date):
        return date_text(value)
    return str(value if value is not None else "").replace("\r", " ").replace("\n", " ")


def read_config(input_dir: Path) -> tuple[dict[str, Any], dict[str, str]]:
    payload = json.loads((input_dir / "run_config.json").read_text(encoding="utf-8"))
    pairs = read_csv(input_dir / "config.csv")
    settings = {as_text(row.get("Key")): row.get("Value", "") for row in pairs}
    return payload, settings


def tax_vendor_map(rows: list[dict[str, str]]) -> dict[str, str]:
    return {normalise(row.get("TaxCode")): as_text(row.get("CanonicalVendor")) for row in rows if normalise(row.get("TaxCode"))}


def profile_by_tax(rows: list[dict[str, str]]) -> dict[str, dict[str, str]]:
    return {normalise(row.get("TaxCode")): row for row in rows if normalise(row.get("TaxCode"))}


def extract_regex(text: str, pattern: str) -> str:
    if not pattern:
        return ""
    try:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    except re.error:
        return ""
    if not match:
        return ""
    return as_text(match.group(1) if match.lastindex else match.group(0))


def clean_invoice_no(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    if len(digits) == 8:
        return digits
    return value.strip()


def extract_invoice_metadata(text: str, profiles: dict[str, dict[str, str]], vendors: dict[str, str]) -> tuple[str, str, date | None, str]:
    tax_code_match = re.search(r"\b\d{10}\b", text)
    tax_code = tax_code_match.group(0) if tax_code_match else ""
    profile = profiles.get(normalise(tax_code), {})
    invoice_no = clean_invoice_no(extract_regex(text, as_text(profile.get("InvoiceNoPattern"))))
    if not invoice_no:
        match = re.search(r"\b0{4}\d{4}\b", text)
        invoice_no = match.group(0) if match else ""
    date_pattern = as_text(profile.get("DatePattern"))
    parsed_date = None
    if date_pattern:
        try:
            match = re.search(date_pattern, text, re.IGNORECASE | re.DOTALL)
        except re.error:
            match = None
        if match and match.lastindex and match.lastindex >= 3:
            try:
                parsed_date = date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
            except ValueError:
                parsed_date = None
    if not parsed_date:
        date_patterns = (
            r"Ng.y\s*\(Date\)\s*([0-3]?\d)\s*th.ng\s*\(month\)\s*([01]?\d)\s*n.m\s*\(year\)\s*(20\d{2})",
            r"([0-3]?\d)\s*\|\s*th.ng\s*\|\s*([01]?\d)\s*\|\s*n.m\s*\|\s*(20\d{2})",
            r"(?:Ng.y|Date)\D{0,20}([0-3]?\d)[/-]([01]?\d)[/-](20\d{2})",
        )
        for pattern in date_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                try:
                    parsed_date = date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
                    break
                except ValueError:
                    continue
    return invoice_no, tax_code, parsed_date, as_text(vendors.get(normalise(tax_code), "UNMAPPED"))


def parse_quantity(raw: str) -> Decimal:
    return decimal_value(raw)


def extract_vat_lines(text: str, profile: dict[str, str]) -> list[tuple[str, str, Decimal]]:
    """Extract material/quantity tuples using profile first, then safe defaults."""
    candidates: list[tuple[str, str, Decimal]] = []
    patterns: list[str] = []
    profile_pattern = as_text(profile.get("LinePattern"))
    if profile_pattern:
        patterns.append(profile_pattern)
    patterns.extend(
        [
            r"Linh ki.n\s+([A-Z0-9]+)[^|]*\|\s*[^|]*\|\s*([0-9.,]+)",
            r"([A-Z]{3}[A-Z0-9]*\d[A-Z0-9]*)[\s\S]{0,220}?\|\s*\d{1,2}\s*\|\s*[^|]+\|\s*([0-9.,]+)",
            r"\|\s*\d{1,2}\s*\|\s*([A-Z]{3}[A-Z0-9]*\d[A-Z0-9]*)[^|]*\|\s*[^|]+\|\s*([0-9.,]+)",
        ]
    )
    # A later fallback pattern can describe the same cell sequence.  Retain
    # duplicates only when they are clearly separate occurrences in the PDF;
    # the same material+qty detected by two parsers is emitted once.
    seen: set[tuple[str, str]] = set()
    for pattern in patterns:
        try:
            matches = list(re.finditer(pattern, text, re.IGNORECASE | re.DOTALL))
        except re.error:
            continue
        for match in matches:
            if len(match.groups()) < 2:
                continue
            material_raw, qty_raw = as_text(match.group(1)), as_text(match.group(2))
            material = normalise(material_raw)
            qty = parse_quantity(qty_raw)
            signature = (material, quantity_text(qty))
            if not material or qty <= 0 or signature in seen:
                continue
            seen.add(signature)
            candidates.append((material_raw, qty_raw, qty))
        if candidates and pattern == profile_pattern:
            break
    return candidates


def fingerprint(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_pdfs(
    pdf_folder: Path,
    profiles: dict[str, dict[str, str]],
    vendors: dict[str, str],
    logs: list[dict[str, Any]],
    run_id: str,
) -> tuple[list[dict[str, Any]], list[Invoice], list[VatLine]]:
    pdf_rows: list[dict[str, Any]] = []
    invoices: list[Invoice] = []
    vat_lines: list[VatLine] = []
    for file_index, pdf_path in enumerate(sorted(pdf_folder.glob("*.pdf")), start=1):
        pdf_id = f"PDF-{file_index:04d}"
        row: dict[str, Any] = {
            "PdfID": pdf_id,
            "OriginalPath": str(pdf_path),
            "Fingerprint": "",
            "ExtractStatus": "PARSE_ERROR",
            "InvoiceCount": 0,
        }
        try:
            row["Fingerprint"] = fingerprint(pdf_path)
            document = pymupdf.open(pdf_path)
            pages = [(page_index + 1, page.get_text("text")) for page_index, page in enumerate(document)]
            document.close()
        except Exception as exc:  # PDF corruption must not stop its batch peers
            logs.append(log_row(run_id, "ERROR", "parse_pdfs", "PDF_OPEN_FAILED", str(exc), str(pdf_path)))
            pdf_rows.append(row)
            continue
        if not any(text.strip() for _, text in pages):
            row["ExtractStatus"] = "NEEDS_OCR"
            logs.append(log_row(run_id, "WARNING", "parse_pdfs", "NEEDS_OCR", "PDF has no text layer.", str(pdf_path)))
            pdf_rows.append(row)
            continue
        fragments: dict[str, dict[str, Any]] = {}
        current_key = ""
        for page_no, text in pages:
            invoice_no, tax_code, invoice_date, vendor = extract_invoice_metadata(text, profiles, vendors)
            if invoice_no:
                current_key = invoice_no
                if current_key not in fragments:
                    fragments[current_key] = {
                        "invoice_no": invoice_no, "tax_code": tax_code, "invoice_date": invoice_date,
                        "vendor": vendor, "pages": [], "texts": [],
                    }
            if current_key:
                fragments[current_key]["pages"].append(page_no)
                fragments[current_key]["texts"].append(text)
        for invoice_index, fragment in enumerate(fragments.values(), start=1):
            invoice_id = f"INV-{fragment['invoice_no']}"
            # Same invoice number in distinct PDFs remains distinguishable in
            # allocations and reports; the display number stays unchanged.
            if any(existing.invoice_id == invoice_id for existing in invoices):
                invoice_id = f"{invoice_id}-{pdf_id}"
            record = Invoice(
                invoice_id=invoice_id,
                pdf_id=pdf_id,
                invoice_no=fragment["invoice_no"],
                invoice_date=fragment["invoice_date"],
                vendor=fragment["vendor"],
                tax_code=fragment["tax_code"],
                page_from=min(fragment["pages"]),
                page_to=max(fragment["pages"]),
            )
            invoices.append(record)
            profile = profiles.get(normalise(record.tax_code), {})
            all_text = "\n".join(fragment["texts"])
            for seq, (material_raw, qty_raw, qty) in enumerate(extract_vat_lines(all_text, profile), start=1):
                vat_lines.append(
                    VatLine(
                        vat_line_id=f"VL-{record.invoice_no}-{invoice_index:02d}-{seq:03d}",
                        invoice_id=record.invoice_id,
                        seq=seq,
                        material_raw=material_raw,
                        material=normalise(material_raw),
                        qty_raw=qty_raw,
                        qty=qty,
                    )
                )
        row["InvoiceCount"] = len(fragments)
        row["ExtractStatus"] = "PARSED" if fragments else "PARSE_ERROR"
        if not fragments:
            logs.append(log_row(run_id, "WARNING", "parse_pdfs", "INVOICE_NOT_FOUND", "No invoice number was parsed.", str(pdf_path)))
        pdf_rows.append(row)
    return pdf_rows, invoices, vat_lines


def log_row(run_id: str, severity: str, procedure: str, code: str, message: str, original_path: str = "", new_path: str = "") -> dict[str, str]:
    return {
        "Timestamp": utc_timestamp(), "RunID": run_id, "Severity": severity,
        "Module": "PythonEngine", "Procedure": procedure, "Code": code,
        "Message": message, "OriginalPath": original_path, "NewPath": new_path,
    }


def parse_gr(rows: list[dict[str, str]], material_vendor: dict[str, str]) -> list[GrRow]:
    parsed: list[GrRow] = []
    for row in rows:
        material = normalise(row.get("Material"))
        vendor = as_text(material_vendor.get(material) or row.get("Vendor"))
        parsed.append(
            GrRow(
                source_row=as_text(row.get("SourceRow")),
                receipt_no=as_text(row.get("ReceiptNo")),
                material=material,
                vendor=vendor,
                qty=decimal_value(row.get("QtyMatch")),
                gr_date=parse_date(row.get("GRDate")),
                flags=as_text(row.get("Flags")).upper(),
            )
        )
    return parsed


def is_eligible(row: GrRow, invoice_date: date | None, settings: Settings, used_sources: set[str]) -> bool:
    if not row.source_row or not row.receipt_no or row.qty <= 0 or row.source_row in used_sources:
        return False
    if "QTY_DOC_ACTUAL_MISMATCH" in row.flags:
        return False
    if invoice_date is None or row.gr_date is None:
        return True
    delta = (row.gr_date - invoice_date).days
    return -settings.gr_before_invoice_days <= delta <= settings.gr_after_invoice_days


def date_distance(row: GrRow, invoice_date: date | None) -> int:
    if not invoice_date or not row.gr_date:
        return 1000
    return abs((row.gr_date - invoice_date).days)


def units_for(value: Decimal) -> int | None:
    scaled = value * Decimal(1000)
    rounded = scaled.to_integral_value(rounding=ROUND_HALF_UP)
    if abs(scaled - rounded) > EPSILON:
        return None
    result = int(rounded)
    return result if result >= 0 else None


def plan_rank(plan: ExactPlan, selected_receipts: set[str], receipt_breadth: Counter[str]) -> tuple[Any, ...]:
    receipts = plan.receipts
    new_receipts = len(receipts - selected_receipts)
    shared_receipts = len(receipts & selected_receipts)
    breadth = sum(receipt_breadth[receipt] for receipt in receipts)
    return (new_receipts, -shared_receipts, -breadth, plan.date_distance, len(plan.choices), tuple(sorted(receipts)))


def plans_for_code(
    rows: list[GrRow],
    demand: Decimal,
    invoice_date: date | None,
    selected_receipts: set[str],
    receipt_breadth: Counter[str],
) -> list[ExactPlan]:
    """Find exact whole-GR-row combinations, retaining a small ranked frontier.

    The bound protects a poorly configured material map from building an
    enormous subset-sum state space.  A bounded-out code is reported for
    review, never silently allocated by a partial/approximate amount.
    """
    target = units_for(demand)
    if target is None or target <= 0:
        return []
    unique: dict[str, GrRow] = {row.source_row: row for row in rows}
    ordered = sorted(
        unique.values(),
        key=lambda row: (
            0 if row.receipt_no in selected_receipts else 1,
            -receipt_breadth[row.receipt_no], date_distance(row, invoice_date),
            row.receipt_no, row.source_row,
        ),
    )[:MAX_PLAN_ROWS]
    states: dict[int, list[tuple[SourceChoice, ...]]] = {0: [tuple()]}
    for row in ordered:
        row_units = units_for(row.qty)
        if row_units is None or row_units <= 0 or row_units > target:
            continue
        choice = SourceChoice(row.source_row, row.receipt_no, row.qty, date_distance(row, invoice_date))
        additions: dict[int, list[tuple[SourceChoice, ...]]] = defaultdict(list)
        for subtotal, existing_plans in list(states.items()):
            new_total = subtotal + row_units
            if new_total > target:
                continue
            for existing in existing_plans:
                additions[new_total].append(existing + (choice,))
        for subtotal, plans in additions.items():
            candidates = states.get(subtotal, []) + plans
            unique_plans: dict[tuple[str, ...], ExactPlan] = {
                tuple(choice.source_row for choice in plan): ExactPlan(plan) for plan in candidates
            }
            states[subtotal] = [
                exact.choices
                for exact in sorted(unique_plans.values(), key=lambda p: plan_rank(p, selected_receipts, receipt_breadth))[:MAX_PLANS_PER_TOTAL]
            ]
        if len(states) > MAX_DP_STATES:
            # States closest to target are most valuable; the score also
            # preserves an empty state so later rows may still start a plan.
            kept = sorted(states, key=lambda subtotal: (abs(target - subtotal), -subtotal))[:MAX_DP_STATES - 1]
            states = {0: states[0], **{subtotal: states[subtotal] for subtotal in kept if subtotal != 0}}
    choices = states.get(target, [])
    unique_results = {tuple(choice.source_row for choice in plan): ExactPlan(plan) for plan in choices}
    return sorted(unique_results.values(), key=lambda p: plan_rank(p, selected_receipts, receipt_breadth))


def allocate_plan_to_vat_lines(
    plan: ExactPlan,
    lines: list[VatLine],
    invoice_id: str,
    allocation_start: int,
) -> tuple[list[dict[str, Any]], int]:
    capacity = [{"choice": choice, "remaining": choice.qty} for choice in plan.choices]
    rows: list[dict[str, Any]] = []
    next_id = allocation_start
    for line in sorted(lines, key=lambda entry: entry.seq):
        remaining = line.qty
        for bucket in capacity:
            if remaining <= 0:
                break
            available = bucket["remaining"]
            if available <= 0:
                continue
            allocated = min(remaining, available)
            bucket["remaining"] -= allocated
            remaining -= allocated
            choice: SourceChoice = bucket["choice"]
            rows.append(
                {
                    "AllocationID": f"A-{next_id:06d}", "InvoiceID": invoice_id,
                    "VatLineID": line.vat_line_id, "ReceiptNo": choice.receipt_no,
                    "SourceRow": choice.source_row, "AllocatedQty": allocated,
                    "Residual": remaining, "Status": "AUTO_MATCHED",
                }
            )
            next_id += 1
    if any(bucket["remaining"] != 0 for bucket in capacity) or remaining != 0:
        raise ValueError("Exact code plan could not be assigned to VAT lines")
    return rows, next_id


def scope_for(material: str, gr_materials: set[str], out_of_scope: set[str]) -> str:
    if material in gr_materials:
        return "IN_SCOPE"
    if material in out_of_scope:
        return "OUT_OF_SCOPE_MATERIAL"
    return "UNKNOWN_MATERIAL"


def receipt_text(receipts: Iterable[str]) -> str:
    def key(value: str) -> tuple[int, Any]:
        return (0, int(value)) if value.isdigit() else (1, value)

    return "+".join(sorted(set(receipts), key=key))


def make_note(matched_count: int, total_count: int, receipts: str, unmatched: list[str], out_scope_count: int, missing_invoice_date: bool) -> str:
    if total_count:
        note = f"Khớp {matched_count}/{total_count} mã với số phiếu {receipts or 'chưa có'}"
    else:
        note = "Không có mã vật tư thuộc phạm vi file P."
    if unmatched:
        shown = ", ".join(unmatched[:12])
        suffix = "" if len(unmatched) <= 12 else f" và {len(unmatched) - 12} mã khác"
        note += f"; còn {len(unmatched)} mã {shown}{suffix} không tìm thấy phiếu phù hợp"
    if out_scope_count:
        note += f"; bỏ qua {out_scope_count} mã đã xác nhận ngoài xưởng"
    if missing_invoice_date:
        note += "; không đọc được ngày hóa đơn nên cần duyệt thủ công"
    return note + "."


def match_batch(
    invoices: list[Invoice],
    vat_lines: list[VatLine],
    gr_rows: list[GrRow],
    out_of_scope: set[str],
    settings: Settings,
    run_id: str,
) -> dict[str, list[dict[str, Any]]]:
    """Perform capacity-safe, many-receipt matching.

    A material's VAT quantity is matched to the *sum* of whole GR rows.  The
    matcher optimises number of completed codes first, then reuses receipts
    that cover other codes, then minimises invoice/GR date distance.
    """
    lines_by_invoice: dict[str, list[VatLine]] = defaultdict(list)
    for line in vat_lines:
        lines_by_invoice[line.invoice_id].append(line)
    gr_by_vendor_material: dict[tuple[str, str], list[GrRow]] = defaultdict(list)
    gr_materials: set[str] = set()
    for row in gr_rows:
        gr_by_vendor_material[(normalise(row.vendor), row.material)].append(row)
        gr_materials.add(row.material)
    for line in vat_lines:
        line.scope_status = scope_for(line.material, gr_materials, out_of_scope)

    allocations: list[dict[str, Any]] = []
    candidates: list[dict[str, Any]] = []
    reports: list[dict[str, Any]] = []
    receipt_report: list[dict[str, Any]] = []
    logs: list[dict[str, Any]] = []
    used_sources: set[str] = set()
    allocation_id = 1
    # Date-first ordering reduces avoidable source ownership conflicts and is
    # deterministic when invoice dates tie.
    ordered_invoices = sorted(invoices, key=lambda inv: (inv.invoice_date or date.max, inv.invoice_no, inv.invoice_id))
    invoice_status: dict[str, str] = {}
    for invoice in ordered_invoices:
        lines = lines_by_invoice.get(invoice.invoice_id, [])
        demand_by_material: dict[str, list[VatLine]] = defaultdict(list)
        out_scope_codes: set[str] = set()
        for line in lines:
            if line.scope_status == "OUT_OF_SCOPE_MATERIAL":
                out_scope_codes.add(line.material)
            else:
                demand_by_material[line.material].append(line)
        if not demand_by_material:
            status = "OTHER_FACTORY" if out_scope_codes else "SUSPECT"
            reports.append(
                report_row(run_id, invoice, "", Decimal(0), status,
                           "Không có mã vật tư thuộc phạm vi file P." if out_scope_codes else "Không trích xuất được mã vật tư để đối soát.")
            )
            invoice_status[invoice.invoice_id] = status
            continue

        eligible_by_material: dict[str, list[GrRow]] = {}
        receipt_breadth: Counter[str] = Counter()
        for material, material_lines in demand_by_material.items():
            candidates_for_material = [
                row for row in gr_by_vendor_material.get((normalise(invoice.vendor), material), [])
                if is_eligible(row, invoice.invoice_date, settings, used_sources)
            ]
            eligible_by_material[material] = candidates_for_material
            for receipt in {row.receipt_no for row in candidates_for_material}:
                receipt_breadth[receipt] += 1
        # Least flexible codes are planned first.  At every step the plan
        # scorer favours receipts that match many invoice codes, exactly as the
        # operational rule requests.
        material_order = sorted(
            demand_by_material,
            key=lambda material: (len({row.receipt_no for row in eligible_by_material[material]}), -sum(line.qty for line in demand_by_material[material]), material),
        )
        selected_receipts: set[str] = set()
        selected_plans: dict[str, ExactPlan] = {}
        unmatched: list[str] = []
        for material in material_order:
            demand = sum((line.qty for line in demand_by_material[material]), Decimal(0))
            plans = plans_for_code(eligible_by_material[material], demand, invoice.invoice_date, selected_receipts, receipt_breadth)
            if not plans:
                unmatched.append(material)
                continue
            plan = plans[0]
            selected_plans[material] = plan
            selected_receipts.update(plan.receipts)
            used_sources.update(choice.source_row for choice in plan.choices)
            added, allocation_id = allocate_plan_to_vat_lines(plan, demand_by_material[material], invoice.invoice_id, allocation_id)
            allocations.extend(added)

        matched_count = len(selected_plans)
        total_count = len(demand_by_material)
        receipts = receipt_text(receipt for plan in selected_plans.values() for receipt in plan.receipts)
        coverage = (Decimal(100) * Decimal(matched_count) / Decimal(total_count)).quantize(Decimal("0.1"))
        missing_date = invoice.invoice_date is None
        if matched_count == total_count and not missing_date:
            status = "MATCHED"
        elif matched_count:
            status = "PARTIAL_MATCHED"
        else:
            status = "SUSPECT"
        note = make_note(matched_count, total_count, receipts, unmatched, len(out_scope_codes), missing_date)
        reports.append(report_row(run_id, invoice, receipts, coverage, status, note))
        invoice_status[invoice.invoice_id] = status
        candidates.append(
            {
                "CandidateID": f"PY-{invoice.invoice_id}", "InvoiceID": invoice.invoice_id,
                "ReceiptSet": receipts, "VendorScore": 20 if invoice.vendor != "UNMAPPED" else 0,
                "MaterialScore": coverage * Decimal("0.4"), "QtyScore": coverage * Decimal("0.2"),
                "DateScore": date_score(selected_plans, invoice.invoice_date), "HintScore": 0,
                "TotalScore": coverage, "Rank": 1,
                "Reasons": f"PYTHON_SUM_QTY: exact quantity for {matched_count}/{total_count} material codes; receipts ranked by multi-code coverage then date proximity.",
            }
        )
        logs.append(log_row(run_id, "INFO", "match_batch", "INVOICE_MATCHED", f"{invoice.invoice_no}: {note}"))

    receipt_invoices: dict[str, set[str]] = defaultdict(set)
    receipt_dates: dict[str, date | None] = {}
    receipt_vendors: dict[str, str] = {}
    for allocation in allocations:
        receipt_invoices[allocation["ReceiptNo"]].add(allocation["InvoiceID"])
    for row in gr_rows:
        receipt_dates.setdefault(row.receipt_no, row.gr_date)
        receipt_vendors.setdefault(row.receipt_no, row.vendor)
    for receipt, linked in sorted(receipt_invoices.items()):
        status = "MATCHED" if all(invoice_status.get(inv) == "MATCHED" for inv in linked) else "REVIEW"
        receipt_report.append(
            {
                "ReceiptNo": receipt, "Vendor": receipt_vendors.get(receipt, ""),
                "GRDate": receipt_dates.get(receipt), "LinkedInvoices": ",".join(sorted(linked)),
                "Coverage": len(linked), "Status": status,
            }
        )
    return {
        "vat_lines": [vat_line_row(line) for line in vat_lines], "match_candidates": candidates,
        "allocations": allocations, "invoice_report": reports, "receipt_report": receipt_report,
        "logs": logs,
    }


def date_score(plans: dict[str, ExactPlan], invoice_date: date | None) -> Decimal:
    if not plans or invoice_date is None:
        return Decimal(0)
    distance = sum(plan.date_distance for plan in plans.values())
    rows = sum(len(plan.choices) for plan in plans.values())
    if not rows:
        return Decimal(0)
    average = Decimal(distance) / Decimal(rows)
    return max(Decimal(0), Decimal(10) - average).quantize(Decimal("0.1"))


def report_row(run_id: str, invoice: Invoice, receipts: str, score: Decimal, status: str, note: str) -> dict[str, Any]:
    return {
        "RunID": run_id, "PDF": invoice.pdf_id, "InvoiceNo": invoice.invoice_no,
        "Vendor": invoice.vendor, "InvoiceDate": invoice.invoice_date,
        "ProposedReceipts": receipts, "TotalScore": score, "Status": status,
        "RenamePreview": "", "Decision": "", "Note": note,
    }


def vat_line_row(line: VatLine) -> dict[str, Any]:
    return {
        "VatLineID": line.vat_line_id, "InvoiceID": line.invoice_id, "Seq": line.seq,
        "MaterialRaw": line.material_raw, "MaterialNorm": line.material, "QtyRaw": line.qty_raw,
        "Qty": line.qty, "Unit": line.unit, "ParseConfidence": line.confidence,
        "ScopeStatus": line.scope_status,
    }


def invoice_row(invoice: Invoice) -> dict[str, Any]:
    return {
        "InvoiceID": invoice.invoice_id, "PdfID": invoice.pdf_id, "PageFrom": invoice.page_from,
        "PageTo": invoice.page_to, "InvoiceNo": invoice.invoice_no, "InvoiceDate": invoice.invoice_date,
        "VendorCanonical": invoice.vendor, "TaxCode": invoice.tax_code, "ParseStatus": invoice.parse_status,
    }


def run(input_dir: Path, output_dir: Path) -> dict[str, Any]:
    payload, config = read_config(input_dir)
    run_id = as_text(payload.get("run_id")) or f"PY-{datetime.now():%Y%m%d-%H%M%S}"
    pdf_folder = Path(as_text(payload.get("pdf_folder")))
    if not pdf_folder.is_dir():
        raise ValueError(f"PDF folder does not exist: {pdf_folder}")
    settings = Settings.from_config(config)
    logs: list[dict[str, Any]] = [
        log_row(run_id, "INFO", "run", "PYTHON_ENGINE_STARTED", "Python PDF parser and many-receipt matcher started.", str(pdf_folder))
    ]
    vendors = tax_vendor_map(read_csv(input_dir / "ncc_map.csv"))
    profiles = profile_by_tax(read_csv(input_dir / "parser_profiles.csv"))
    material_vendor = {
        normalise(row.get("MaterialNorm")): as_text(row.get("CanonicalVendor"))
        for row in read_csv(input_dir / "material_ncc_map.csv")
        if normalise(row.get("MaterialNorm")) and as_text(row.get("CanonicalVendor"))
    }
    out_of_scope = {
        normalise(row.get("MaterialNorm"))
        for row in read_csv(input_dir / "material_scope_map.csv")
        if as_text(row.get("ScopeStatus")).upper() == "OUT_OF_SCOPE_MATERIAL"
    }
    pdf_rows, invoices, vat_lines = parse_pdfs(pdf_folder, profiles, vendors, logs, run_id)
    gr_rows = parse_gr(read_csv(input_dir / "gr_data.csv"), material_vendor)
    matched = match_batch(invoices, vat_lines, gr_rows, out_of_scope, settings, run_id)
    logs.extend(matched.pop("logs"))
    logs.append(log_row(run_id, "INFO", "run", "PYTHON_ENGINE_COMPLETED", f"Parsed {len(invoices)} invoices; created {len(matched['invoice_report'])} report rows."))
    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(output_dir / "pdf_files.csv", PDF_FIELDS, pdf_rows)
    write_csv(output_dir / "invoices.csv", INVOICE_FIELDS, [invoice_row(invoice) for invoice in invoices])
    write_csv(output_dir / "vat_lines.csv", VAT_LINE_FIELDS, matched["vat_lines"])
    write_csv(output_dir / "match_candidates.csv", CANDIDATE_FIELDS, matched["match_candidates"])
    write_csv(output_dir / "allocations.csv", ALLOCATION_FIELDS, matched["allocations"])
    write_csv(output_dir / "invoice_report.csv", REPORT_FIELDS, matched["invoice_report"])
    write_csv(output_dir / "receipt_report.csv", RECEIPT_REPORT_FIELDS, matched["receipt_report"])
    write_csv(output_dir / "engine_log.csv", LOG_FIELDS, logs)
    summary = {
        "run_id": run_id, "invoice_count": len(invoices), "pdf_count": len(pdf_rows),
        "report_count": len(matched["invoice_report"]), "allocation_count": len(matched["allocations"]),
        "settings": {"GRBeforeInvoiceDays": settings.gr_before_invoice_days, "GRAfterInvoiceDays": settings.gr_after_invoice_days},
    }
    (output_dir / "engine_result.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="VAT Matcher PDF parser and matcher")
    parser.add_argument("--input", required=True, type=Path, help="Excel-exported input directory")
    parser.add_argument("--output", required=True, type=Path, help="Directory for normalised output CSVs")
    args = parser.parse_args(argv)
    try:
        summary = run(args.input, args.output)
    except Exception as exc:
        print(f"VAT Matcher Python engine failed: {exc}", file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
