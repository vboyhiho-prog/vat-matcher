# VAT Matcher

## What it does

This Excel VBA workbook matches VAT invoices to GR receipts without installing `pdftotext` or another external PDF program. Excel Power Query reads all PDFs in a selected folder through `Pdf.Tables`.

## Daily workflow

1. Open the workbook and enable macros.
2. The release workbook is already initialized. Do not run setup for normal daily use.
3. Run `PickAndLoadTrackingFile` and select the GR tracking workbook.
4. Run `ConfigurePdfFolderBatch`, choose the PDF folder, then run `RefreshPdfBatch`.
5. Run `ParsePdfRawToInvoices`, `ParseVatLinesFromPdfRaw`, then `RunMatch`.
6. Review `BC_HOA_DON`, `MATCH_CANDIDATES`, `ALLOCATIONS`, and `LOG`.
7. If needed, enter a row in `MANUAL_OVERRIDES` using `APPROVE` or `REJECT`, then run `ApplyManualOverrides`.
8. Run `RunWorkbookSelfTest` after an upgrade or before a pilot batch. The `TEST_RESULTS` sheet should show only `PASS`.

## Matching safeguards

- Candidate score is out of 100: vendor 20, material coverage 40, quantity 20, date 10/8/5, and a validated email-IB hint 10.
- A candidate more than two days from the GR date is rejected. A `HIGH` email hint only adds points when its IB is present in `GR_DATA` for the proposed receipt.
- `PARTIAL_MATCHED` is a review result: it keeps the proposed receipt set and filename, while its score is the percentage of in-scope VAT lines matched exactly. Repeated material codes are scored line by line because their quantities may belong to different receipts. It is never treated as a 100% match.
- Allocation is capacity-safe across the whole batch. A GR `SourceRow` can never be allocated above `QtyMatch` or assigned to more than one invoice; any residual or conflict becomes `SUSPECT_CONFLICT` for review.
- Rows flagged `QTY_DOC_ACTUAL_MISMATCH` are review-only and excluded from automatic allocation by default.
- Search is capped at 5,000 receipt candidates per invoice. If exceeded, the result is `SUSPECT` with `SEARCH_TRUNCATED`; it is never auto-matched.

## MSG email intake

Run `ExtractPdfAttachmentsFromMsg` and select an `.msg` file. Then select an ASCII-only output folder such as `C:\VAT_Matching_Work`. The macro saves only `.pdf` attachments and records skipped image attachments in `EMAIL_ATTACHMENTS` and `LOG`. It does not send, reply to, or modify email.

## Add a new supplier

Use `NCC_CONFIG_IMPORT`. Replace the `EXAMPLE_VENDOR` row with the supplier name, tax code, aliases, profile ID, and initial patterns. Set `ProfileStatus` to `DRAFT`, then run `LoadVendorProfilesFromConfig`. The workbook creates or updates the supplier in `NCC_MAP` and its profile in `PARSER_PROFILES`.

Keep a profile as `DRAFT` until you have tested at least one representative PDF and reviewed `PQ_PDF_RAW`, `INVOICES`, and `VAT_LINES`. Change it to `ACTIVE` only after the test is accepted. A new configuration does not need a separate worksheet or a separate workbook.

## Invoices that include another factory's materials

`VAT_LINES.ScopeStatus` shows how each parsed material is handled:

- `IN_SCOPE`: the code exists in the current `GR_DATA` snapshot and is matched normally.
- `OUT_OF_SCOPE_MATERIAL`: a code you have explicitly entered in `MATERIAL_SCOPE_MAP` as outside your factory. It is kept for audit but excluded from allocation and score coverage.
- `UNKNOWN_MATERIAL`: not in GR and not yet confirmed external. It is not ignored; review it so missing internal data is never silently treated as another factory.

To classify a confirmed external code, add its normalized code to `MATERIAL_SCOPE_MAP`, set `ScopeStatus` to `OUT_OF_SCOPE_MATERIAL`, and add a short note. An invoice whose every line is confirmed external becomes `OTHER_FACTORY`. Its rename preview is `VAT XUONG KHAC.pdf`; apply it only through the normal copied-work-folder rename step, so collision checks and rollback still apply.

## Rename safety

`CreateRenamePreviews` creates a preview only. It never changes a filename. After reviewing the proposed receipts and filename, type `OK` manually in the `Decision` column in `BC_HOA_DON`, run `ApplyApprovedRenames`, and select the folder containing the PDF. `OK` is the user's explicit file-level authorization, including a partial or multi-invoice proposal; blank/`NG` rows are not renamed. Existing target files are never overwritten. Run `RollbackRenames` with the same folder to reverse logged rename actions.

## Sending results back for support

After a company run, use `ExportDiagnosticBundle` and choose a writable folder. It creates a macro-free `.xlsx` containing `LOG`, `BC_HOA_DON`, `BC_PHIEU`, `VAT_LINES`, `PDF_FILES`, `EMAIL_HINTS`, `MATERIAL_SCOPE_MAP`, `CONFIG`, and `TEST_RESULTS`. Send that file together with the exact PDF filename that needs investigation. The export does not rename, delete, or alter your source PDFs.

## Limitations

- The parser is validated against the supplied invoice layout. A supplier layout that differs materially may need a parser profile.
- A scanned PDF or a page with no text is recorded as `NEEDS_OCR`; the batch continues, but this v1 tool deliberately does not install or run OCR.
- `SUSPECT` and `PARTIAL_MATCHED` need human review. A proposed name may still be shown, but rename requires a manually entered `OK`.
- Power Query PDF requires an Excel edition that includes Power Query's PDF connector.
