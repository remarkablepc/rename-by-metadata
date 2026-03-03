# Pause before exit to allow error review
Read-Host "Press Enter to exit"
# build-Rename-ByMetadata.ps1 - Builds Rename-ByMetadata.exe from Rename-ByMetadata.ps1

$root = $PSScriptRoot
$src  = Join-Path $root "Rename-ByMetadata.ps1"
$out  = Join-Path $root "Rename-ByMetadata.exe"
$icon = Join-Path (Split-Path $root -Parent) "script.ico"

Write-Host "Building Rename-ByMetadata.exe..." -ForegroundColor Cyan

ps2exe `
  -inputFile $src `
  -outputFile $out `
  -iconFile $icon `
  -title "Rename By Metadata" `
  -product "Rename By Metadata" `
  -company "Local Scripts" `
  -description "Metadata-driven preview-first renaming utility" `
  -version "1.0.0.0" `
  -noConsole:$false `
  -x64

Write-Host "EXE build complete: $out"
Write-Host "Checking for code-signing certificate..." -ForegroundColor Yellow

$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1

if ($cert) {
    Write-Host "Signing EXE with certificate: $($cert.Subject)" -ForegroundColor Cyan
    Set-AuthenticodeSignature -FilePath $out -Certificate $cert | Out-Null
    Write-Host "EXE successfully signed." -ForegroundColor Green
} else {
    Write-Host "No code-signing certificate found. Skipping signing step." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Build process finished."
Write-Host "Note: After building the EXE, publish the latest script and changelog to GitHub for release tracking."
