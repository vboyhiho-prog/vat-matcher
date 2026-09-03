@echo off
setlocal EnableExtensions

rem Set VAT_MATCHER_PYTHON once if the company-approved Python is not on PATH.
if defined VAT_MATCHER_PYTHON (
  set "VAT_MATCHER_PYTHON_EXE=%VAT_MATCHER_PYTHON%"
) else (
  set "VAT_MATCHER_PYTHON_EXE=python"
)

"%VAT_MATCHER_PYTHON_EXE%" -c "import pymupdf" >nul 2>nul
if errorlevel 1 (
  echo ERROR: Approved Python with PyMuPDF is unavailable. Set VAT_MATCHER_PYTHON to the fixed tool environment. 1>&2
  exit /b 2
)

"%VAT_MATCHER_PYTHON_EXE%" "%~dp0engine.py" %*
exit /b %errorlevel%
