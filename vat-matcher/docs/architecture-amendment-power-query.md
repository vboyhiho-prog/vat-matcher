# Architecture amendment - batch PDF ingestion with Power Query

Status: approved by the user on 2026-08-31.

## Superseded decision

The original v1 plan required an externally installed, approved `pdftotext.exe`. That dependency is removed. The tool must not download, install, embed, or invoke `pdftotext`.

## Revised input contract

1. A user places 50-100 text-layer invoice PDFs in one work folder.
2. Excel Power Query reads that folder using the built-in `Pdf.Tables` connector and appends recognized invoice tables to a staging query/table.
3. VBA validates the Power Query output and transforms only the required normalized fields into `INVOICES` and `VAT_LINES` before matching.
4. PDF parsing remains report-only when the PDF connector returns no usable table, an error, or a scan. Such files are marked `NEEDS_OCR`/`PARSE_ERROR`; VBA must continue with other files.

## Capability gate

The `HOME` sheet exposes a connector health check. It must confirm that the workbook has the `Pdf.Tables` Power Query function and that a folder refresh can produce data. If unavailable, the tool must state `POWER_QUERY_PDF_UNAVAILABLE`; it must not silently fall back to manual per-file copy/paste.

## Compatibility and limitations

- Requires a Windows Excel edition that includes the Power Query PDF connector. It is not available in all Excel 2016 editions; Excel version 16.0 alone is insufficient proof.
- Requires text-based PDFs with tables recognizable by the connector. OCR/scanned PDFs remain outside v1.
- Foxit Reader is only the user's viewing tool and is not automated by VBA. Foxit PDF Editor batch features are not required.
- The final acceptance test must run on the target company Excel with the supplied sample folder before pilot rollout.

## Unchanged controls

Matching stays deterministic and many-to-many; the source tracking workbook remains read-only; rename requires manual confirmation, preview, collision checks, and rollback log; source samples remain untouched.

## Operational logging requirement

`LOG` is a mandatory visible worksheet backed by `tblLog`, not a developer-only diagnostic. Every public workflow must append a structured record with timestamp, RunID, severity, module, procedure, code, user-facing message, technical detail/recovery action, and only the minimal relevant file paths. `ERROR` and `FATAL` records must remain available after the macro ends so a later Codex session can inspect the workbook, identify the failing procedure, and implement a targeted fix. Log entries must never include full email bodies or other unnecessary sensitive content.
