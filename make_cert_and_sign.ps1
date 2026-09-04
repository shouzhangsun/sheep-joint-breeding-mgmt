# Generate self-signed code-signing cert, export public key, Authenticode-sign
# the exe with PowerShell's built-in Set-AuthenticodeSignature (no Windows SDK needed),
# and write cert_thumb.iss (cert thumbprint) for the Inno Setup uninstall cleanup.
$ErrorActionPreference = "Stop"

# ---- Publisher / version config ----
# $Publisher : shown in installer / Programs and Features / EXE version info (Chinese OK).
#              Keep in sync with installer.iss MyAppPublisher.
# $Version   : keep in sync with installer.iss MyAppVersion.
# $CertSubject : cert Subject CN. MUST be ASCII. The Chinese publisher is presented via
#                the installer and EXE version info instead.
$Publisher   = "QuPinGaiZhan-shouzhangsun"
$Version     = "3.0.0"
$CertSubject = "CN=SheepBreeding Code Signing"
# --------------------------------------

$pw = "SheepBreeding2026"
if ($PSScriptRoot) { $scriptDir = $PSScriptRoot } else { $scriptDir = (Get-Location).Path }
$resources = Join-Path $scriptDir "resources"
New-Item -ItemType Directory -Force -Path $resources | Out-Null

# Remove stale / junk certs so we always end up with exactly one clean cert.
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $CertSubject -or $_.Subject -eq "CN=" -or $_.Subject -eq "CN=Code Signing" } | ForEach-Object { Remove-Item "Cert:\CurrentUser\My\$($_.Thumbprint)" }

# 1) Create the code-signing cert.
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $CertSubject -FriendlyName "SheepBreeding Code Signing $Version" -CertStoreLocation Cert:\CurrentUser\My -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(10)
Write-Host "[ok] created cert $($cert.Thumbprint) subject=$($cert.Subject)"

# 2) Export public .cer (installer imports it into Trusted Root on the target machine).
Export-Certificate -Cert $cert -FilePath (Join-Path $resources "SheepBreeding.cer") | Out-Null
Write-Host "[ok] exported public cert"

# 3) Authenticode-sign the exe (PowerShell built-in, no signtool needed).
$exe = Join-Path $scriptDir "dist\SheepBreeding.exe"
if (Test-Path $exe) {
    $signed = Set-AuthenticodeSignature -FilePath $exe -Certificate $cert -TimestampServer "http://timestamp.digicert.com" -HashAlgorithm SHA256 -Force
    if ($signed.Status -eq "Valid") {
        Write-Host "[ok] signed $exe (Status=$($signed.Status))"
    } else {
        Write-Warning "sign status: $($signed.Status) / $($signed.StatusMessage)"
    }
} else {
    Write-Warning "exe not found: $exe ; run pyinstaller build.spec first"
}

# 4) Write thumbprint for installer.iss uninstall cleanup.
Set-Content -Path (Join-Path $scriptDir "cert_thumb.iss") -Value "#define CertThumbprint `"$($cert.Thumbprint)`"" -Encoding ASCII
Write-Host "[ok] cert thumbprint: $($cert.Thumbprint)"
