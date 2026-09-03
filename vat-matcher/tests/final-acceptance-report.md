# VAT Matcher v1.7 — acceptance report

Date: 2026-09-01

## Release verification

| Gate | Evidence | Result |
| --- | --- | --- |
| Workbook reopen and VBA execution | `VAT_Matcher_v1.7.xlsm` reopened through Excel automation; parser, matcher and self-test executed | PASS |
| Workbook self-test | `TEST_RESULTS` contains 14 rows, all `PASS` | PASS |
| Sample PDF parsing | `VATNCC.pdf` produced 5 invoices: 00002966, 00002958, 00002960, 00002961, 00002965 | PASS |
| Ground-truth matching | 2966→5003; 2958/2960/2965→5013; 2961→SUSPECT | PASS |
| LTV reviewable reference | 00002212→4983+5005+5007, `PARTIAL_MATCHED`, scored and left for manual review | PASS |
| MSG intake | 4 PDFs extracted, 2 PNG skipped; rerun does not overwrite | PASS |
| Email IB validation | 185068322–185068328 map to GR receipt 5007 and give a validated hint bonus | PASS |
| Date/capacity/search safety | T14, T20, T21 and T22 PASS; no GR SourceRow shared across invoices | PASS |
| Rename/rollback | Preview-only safeguards, copy-folder rename and rollback test passed | PASS |
| Scope handling | Mixed scope, other-factory classification and other-factory preview fixtures passed | PASS |
| Source integrity | SHA-256 checks of the three supplied sources match their baseline values | PASS |

## Release contents

- `VAT_Matcher_v1.7.xlsm`: runnable workbook with scope classification safeguards.
- `src/modules/*.bas`: exported VBA source.
- `README.md`: daily workflow, supplier onboarding and limitations.
- `tests/phase-2-4-test-report.md`: detailed evidence.

## Pilot gate

The implementation is ready for a supervised pilot when a real batch becomes available. Do not use `C:\VAT_Matching_Test` as the rename work folder. Review all `SUSPECT`, `SUSPECT_CONFLICT`, `NEEDS_OCR`, `SEARCH_TRUNCATED`, and `UNKNOWN_MATERIAL` rows before any manual approval or rename.

## Known limitations

- No OCR is installed or invoked. Scanned PDF text is safely marked `NEEDS_OCR`.
- Supplier layouts remain configuration-led. Keep a new supplier profile in `DRAFT` until reviewed against a representative PDF.
- Mark a code `OUT_OF_SCOPE_MATERIAL` only after it is confirmed external in `MATERIAL_SCOPE_MAP`; unknown codes deliberately remain reviewable.
- The LTV multi-receipt result is a GR-validated, reviewable `MATCH_HINTS` configuration; it is not represented as an unconstrained inference.
