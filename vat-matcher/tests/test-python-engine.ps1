$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
python (Join-Path $root 'tests\test_python_engine.py')
