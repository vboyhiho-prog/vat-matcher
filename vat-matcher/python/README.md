# VAT Matcher Python runtime

`run_tool.bat` is the stable entry point used by the Excel workbook. It takes
`--input <folder>` and `--output <folder>`. Excel creates those folders and
exchanges UTF-8 CSV files only; it does not pass workbook credentials or edit
the source Tracking P file.

The approved internal Python runtime must include `PyMuPDF==1.28.2`. Set the
machine environment variable `VAT_MATCHER_PYTHON` to its `python.exe` when it
is not the `python` command on `PATH`. The batch file does not install packages
or download anything.

The engine never renames files. It emits suggested names and match evidence;
the existing Excel `OK` / `NG` review and rename/rollback controls remain the
only rename authority.
