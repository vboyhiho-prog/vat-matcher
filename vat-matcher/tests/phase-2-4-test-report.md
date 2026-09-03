# VAT Matcher test report - PDF ingestion and matching

Date: 2026-09-01

## Evidence

| Check | Expected | Actual | Result |
| --- | --- | --- | --- |
| Power Query PDF refresh | Read the sample PDF without an external executable | 25 PDF page/table blocks | PASS |
| Invoice extraction | 5 invoices from `VATNCC.pdf` | `00002966`, `00002958`, `00002960`, `00002961`, `00002965` | PASS |
| VAT-line extraction | Sample material and quantity preserved | `BIW32026240AA`, quantity `752` | PASS |
| Tracking import | Read-only import of the sample tracking workbook | 51,272 GR rows | PASS |
| Matching | `00002966` | `5003`, `MATCHED` | PASS |
| Matching | `00002958`, `00002960`, `00002965` | `5013`, `MATCHED` | PASS |
| Matching | `00002961` | `SUSPECT` | PASS |
| MSG intake | 4 PDF attachments extracted and 2 PNG attachments skipped | 4 PDF, 2 PNG | PASS |
| Manual override | An approved receipt override updates only the report and is logged | `00002961 -> 5013` test override | PASS |
| Override rollback | Removing the test override restores the calculated status | `00002961 -> SUSPECT` | PASS |
| Rename preview | A multi-invoice PDF is never eligible for automatic rename | 5 `REVIEW_ONLY` previews | PASS |
| Rename and rollback | Rename only in a copied test folder, then restore the original name | PASS | PASS |
| LTV reference | `00002212` keeps the configured receipt set as a scored review suggestion | `4983+5005+5007`, `PARTIAL_MATCHED`, Decision blank | PASS |
| Supplier configuration | A DRAFT row loads into NCC_MAP and PARSER_PROFILES | PASS | PASS |
| MSG rerun | Repeat extraction to same output folder | 4 `SKIPPED_DUPLICATE`, no overwrite | PASS |
| Email IB hints | Seven known IB values present in GR_DATA create HIGH hints | 7 rows linked to receipt 5007 | PASS |
| Receipt report | Allocation report contains expected receipt rows | 5003 and 5013 | PASS |
| Workbook self-test | Required tables/report checks | 7/7 PASS | PASS |
| Date scoring (T14) | 0/1/2/3 days = 10/8/5/0 | 10/8/5/0 | PASS |
| Email hint bonus (T13) | HIGH IB hint for receipt 5007 adds 10 points | 10 points | PASS |
| Capacity rule (T21) | Incomplete MATCHED allocation is not retained as matched | `MATCHED -> SUSPECT_CONFLICT`; existing `SUSPECT` unchanged | PASS |
| Source ownership (T22) | One GR `SourceRow` is never assigned to two invoices | 0 shared SourceRows | PASS |
| Allocation safety sample regression | Existing five-invoice sample results remain valid after capacity check | 2966→5003; 2958/2960/2965→5013; 2961→SUSPECT | PASS |
| Search limit (T20) | Candidate search stops safely beyond the configured ceiling | 5,000 allowed; 5,001 triggers `SEARCH_TRUNCATED` | PASS |
| Empty PDF text (T18) | Blank page is not treated as a parsed invoice | `NEEDS_OCR`; subsequent pages continue | PASS |
| Parser regression after OCR guard | Sample `VATNCC.pdf` remains unchanged in behavior | 5 invoices and all ground-truth matches | PASS |

## Safety

The three files under `C:\VAT_Matching_Test` were used read-only. No rename, delete, email action, registry change, or external dependency installation was performed in this phase.

## Current scope

The workbook produces PDF raw data, invoices, VAT lines, GR data, candidates, allocations, invoice report, manual overrides, email attachment intake, controlled rename/rollback, and structured LOG entries. Final delivery packaging remains subsequent work.

## Phase 4 hardening note

Candidate scoring now reports a 100-point scale: vendor 20, material coverage 40, quantity 20, date 10/8/5, and validated email-IB hint 10. The email hint index is built once per matching run from `EMAIL_HINTS` and `GR_DATA`; it does not rescan the 51,272 GR rows for every candidate.

Candidate search is explicitly bounded at 5,000 receipts per invoice. A larger candidate set is not auto-matched: it is reported as `SUSPECT` with reason `SEARCH_TRUNCATED`, so the user can narrow the vendor/date data or use an approved manual decision.

## Scope classification hardening (v1.7)

| Check | Expected | Actual | Result |
| --- | --- | --- | --- |
| T26 mixed scope | Confirmed external material is excluded but an internal material remains matchable | `MIXED_SCOPE` reason retained even after capacity pass | PASS |
| T27 other factory | Invoice with only confirmed external materials has no candidate or allocation | `OTHER_FACTORY`, zero allocation | PASS |
| T28 other-factory preview | 100%-external PDF receives a controlled alternate name | `VAT XUONG KHAC.pdf` preview | PASS |
| Sample regression | Existing test materials are not silently treated as external | 2966→5003; 2958/2960/2965→5013; 2961→SUSPECT | PASS |

`MATCH_HINTS` contains the verified LTV reference as a reviewable configuration row. It is not an unreviewed automatic inference; every receipt in the configured set must exist in GR_DATA before it can be applied. The displayed score is the percentage of in-scope VAT lines that can be allocated exactly to the suggested receipts after excluding quantity-mismatch GR rows and quantities already consumed by earlier invoices.
