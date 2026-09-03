# VAT Matcher engine source (development only)

`engine.py` and `run_tool.bat` support development tests and packaging only.
They are not distributed as the production runtime. The workbook in a release
calls `engine\VAT_Matcher_Engine.exe`; Excel creates the CSV folders and does
not pass workbook credentials or edit the source Tracking P file.

The build computer needs `PyMuPDF==1.28.2` and PyInstaller. PyInstaller embeds
that fixed runtime into the release `engine` folder. End-user machines do not
need Python, PyMuPDF, PyInstaller or `VAT_MATCHER_PYTHON`.

The engine never renames files. It emits suggested names and match evidence;
the existing Excel `OK` / `NG` review and rename/rollback controls remain the
only rename authority.
