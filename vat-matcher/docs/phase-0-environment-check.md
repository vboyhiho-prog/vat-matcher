# Phase 0 - Environment check

Checked: 2026-08-31 (Asia/Saigon)

| Check | Result | Evidence |
|---|---|---|
| Workspace Git state | PASS | Local repository on `master`; initial user content is untracked `outputs/`. No commit or remote was created. |
| Workspace writable | PASS | Hermes Builder state was initialized locally. |
| Excel automation | PASS | Excel COM 16.0, 64-bit. |
| Outlook automation | PASS | Outlook COM 16.0.0.20326. |
| VBA project access | PASS | VBIDE was accessed successfully without changing Trust Center or registry. |
| Tracking workbook discovery | PASS | `C:\VAT_Matching_Test\260807_THEO DOI NHAP P.xlsx` opens read-only. `DATA` has required headers and last key row 51273. |
| MSG discovery | PASS | `Re_ KHGH OS NGÀY 27_8.msg` opens through Outlook COM; subject is `Re: KHGH OS NGÀY 27.8`; six attachments comprise four PDFs and two PNGs. |
| VAT PDF readability | PASS (diagnostic only) | `VATNCC.pdf` has five pages with non-empty text layer according to the bundled Python reader. |
| Approved PDF-to-text executable | **BLOCKED** | No `pdftotext.exe` was found on PATH, under Program Files, or in the supplied bundled runtime. The v1 plan requires a user/organization-approved executable and prohibits automatic installation. |

## Locked source fingerprints

| Source | SHA-256 |
|---|---|
| `C:\VAT_Matching_Test\260807_THEO DOI NHAP P.xlsx` | `AC90C3F0C03FE5EA768A30B9F7C55E9EF9A3039A06F91A49C1193F7C08C9C77F` |
| `C:\VAT_Matching_Test\Re_ KHGH OS NGÀY 27_8.msg` | `8AB8E224A8355F1841C11468ABC278BB2ED2ED9F8BEF06C1CA1E7FAB97E2301E` |
| `C:\VAT_Matching_Test\VATNCC.pdf` | `A39870240A570B521A834FF9E7C65BD19A56E749EEDFC9CE354CCA39954BB09C` |

No source file has been copied, renamed, opened for write, or otherwise modified. The same three hashes were reconfirmed after Phase 0 checks.

## Decision

Do not start Phase 1 implementation or create a production `.xlsm` until an approved, licensed `pdftotext.exe` is supplied with its absolute path and version. This prevents creating a workbook that cannot meet its mandatory PDF parsing and end-to-end acceptance gates. No dependency, Office security setting, registry entry, or external system was changed.
