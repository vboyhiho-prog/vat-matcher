# Architecture amendment — Python PDF and matching engine

Status: adopted for v1.42; portable delivery adopted for v1.44 on 2026-09-04.

## Decision

Power Query `Pdf.Tables` is removed from the production PDF path. It varies by
Office installation, is weak for text/scan detection and makes supplier parser
regression testing impractical. Excel remains the configuration, review and
rename interface; Python owns PDF extraction and matching.

```text
Excel xlsm -- CSV input --> engine/VAT_Matcher_Engine.exe -- CSV output --> Excel xlsm
```

The workbook exports `CONFIG`, `GR_DATA`, NCC maps, parser profiles and scope
maps to a unique `vat_python_runtime/<RunID>/input` folder. It waits for the
engine to finish and imports only the normalised output tables. No external
system is contacted. Neither VBA nor Python modifies the Tracking P source.

## Runtime contract

- Entry point: `engine\VAT_Matcher_Engine.exe` beside the release workbook.
- Dependency: the engine is a PyInstaller one-folder bundle with
  `PyMuPDF==1.28.2` and its own embedded Python runtime. Keep `_internal`
  intact beside the executable.
- End-user machines do not use `python` on `PATH`, `pip` or
  `VAT_MATCHER_PYTHON`. Source `python/run_tool.bat` remains development-only.
- Textless PDFs become `NEEDS_OCR`. OCR is an explicit later step, never a
  silent substitution.
- Parser profiles remain governed by `NCC_CONFIG_IMPORT` / `PARSER_PROFILES`;
  a new profile stays `DRAFT` until its PDF sample is accepted.

## Matching contract

1. Group VAT demand by material code and use the sum of all its VAT lines.
2. Find a combination of whole, eligible GR source rows whose sum exactly
   equals that demand; a code may span multiple receipts.
3. Maximise completed material codes, then favour receipt sets that cover more
   invoice codes, then minimise GR/invoice date distance.
4. Permit `GRDate - InvoiceDate` from `-GRBeforeInvoiceDays` through
   `+GRAfterInvoiceDays` (default `-5` through `+2`).
5. Reserve every selected `SourceRow` for one invoice only across the batch.
   A code without an exact plan is reported, not guessed.

`BC_HOA_DON.Note` expresses code coverage and names unmatched codes. The
existing Excel manual decision is unchanged: only a user-entered `OK` permits
the separate rename routine; the Python engine never renames files.
