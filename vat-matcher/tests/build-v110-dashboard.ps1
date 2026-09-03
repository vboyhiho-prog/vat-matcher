$ErrorActionPreference = 'Stop'
Write-Warning 'build-v110-dashboard.ps1 is deprecated; forwarding to build-release.ps1.'
& (Join-Path $PSScriptRoot 'build-release.ps1') @args
