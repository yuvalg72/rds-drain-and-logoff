#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Builds rds-drain-and-logoff.exe from the PowerShell source and signs it with Authenticode.

.DESCRIPTION
    Supports three certificate sources:
      1. Thumbprint  – use an existing certificate already in your cert store
      2. InternalCA  – request a certificate from your enterprise (ADCS) CA
      3. SelfSigned  – create a self-signed certificate (useful for testing;
                       must be trusted via Group Policy or manually on each machine)

.PARAMETER CertSource
    "Thumbprint" | "InternalCA" | "SelfSigned"   (default: Thumbprint)

.PARAMETER Thumbprint
    Thumbprint of an existing code-signing cert in Cert:\CurrentUser\My or
    Cert:\LocalMachine\My.  Required when CertSource = Thumbprint.

.PARAMETER CATemplate
    Certificate template name on your ADCS server.
    Required when CertSource = InternalCA.  Typically "CodeSigning".

.PARAMETER CertSubject
    Subject CN for SelfSigned or InternalCA certificates.
    Default: "RDS Drain and Logoff Tool"

.PARAMETER TimestampServer
    RFC 3161 timestamp server URL.
    Default: DigiCert public TSA (works without a paid certificate).

.EXAMPLE
    # Re-use an existing cert already in your store
    .\build-and-sign.ps1 -CertSource Thumbprint -Thumbprint "AABBCC..."

.EXAMPLE
    # Issue a cert from your internal Active Directory Certificate Services
    .\build-and-sign.ps1 -CertSource InternalCA -CATemplate "CodeSigning"

.EXAMPLE
    # Quick self-signed build (testing / internal lab only)
    .\build-and-sign.ps1 -CertSource SelfSigned
#>

param(
    [ValidateSet("Thumbprint", "InternalCA", "SelfSigned")]
    [string] $CertSource    = "SelfSigned",

    [string] $Thumbprint    = "",
    [string] $CATemplate    = "CodeSigning",
    [string] $CertSubject   = "RDS Drain and Logoff Tool",
    [string] $TimestampServer = "http://timestamp.digicert.com"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir  = $PSScriptRoot
$SourcePs1  = Join-Path $ScriptDir "drain-rds-farm.ps1"
$OutputExe  = Join-Path $ScriptDir "rds-drain-and-logoff.exe"

# ── 1. Resolve the signing certificate ────────────────────────────────────────

function Get-SigningCert {
    switch ($CertSource) {

        "Thumbprint" {
            if (-not $Thumbprint) {
                throw "Provide -Thumbprint when using CertSource=Thumbprint."
            }
            $cert = (Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue) |
                    Where-Object { $_.Thumbprint -eq $Thumbprint } |
                    Select-Object -First 1
            if (-not $cert) {
                throw "Certificate with thumbprint '$Thumbprint' not found in any local store."
            }
            return $cert
        }

        "InternalCA" {
            Write-Host "Requesting certificate from internal CA (template: $CATemplate)..."
            # Get-Certificate submits a CSR to ADCS and retrieves the signed cert automatically.
            $request = Get-Certificate `
                -Template        $CATemplate `
                -SubjectName     "CN=$CertSubject" `
                -CertStoreLocation Cert:\CurrentUser\My
            return $request.Certificate
        }

        "SelfSigned" {
            Write-Warning "Self-signed certificate created. You must add it to 'Trusted Publishers' and 'Trusted Root CA' stores on every machine that will run the EXE."
            $cert = New-SelfSignedCertificate `
                -Type            CodeSigningCert `
                -Subject         "CN=$CertSubject" `
                -HashAlgorithm   SHA256 `
                -KeyExportPolicy Exportable `
                -CertStoreLocation Cert:\CurrentUser\My `
                -NotAfter        (Get-Date).AddYears(3)

            # Trust the cert locally so Defender accepts it on this machine immediately.
            $rootStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
                "Root", "LocalMachine")
            $rootStore.Open("ReadWrite")
            $rootStore.Add($cert)
            $rootStore.Close()

            $pubStore = [System.Security.Cryptography.X509Certificates.X509Store]::new(
                "TrustedPublisher", "LocalMachine")
            $pubStore.Open("ReadWrite")
            $pubStore.Add($cert)
            $pubStore.Close()

            Write-Host "Self-signed cert added to local Trusted Root CA and Trusted Publishers stores."
            return $cert
        }
    }
}

$cert = Get-SigningCert
Write-Host "Using certificate: Subject=$($cert.Subject)  Thumbprint=$($cert.Thumbprint)"

# ── 2. Sign the PowerShell source script ──────────────────────────────────────

Write-Host "Signing $SourcePs1 ..."
$sigResult = Set-AuthenticodeSignature `
    -FilePath        $SourcePs1 `
    -Certificate     $cert `
    -TimestampServer $TimestampServer `
    -HashAlgorithm   SHA256

if ($sigResult.Status -ne "Valid") {
    throw "Script signing failed: $($sigResult.StatusMessage)"
}
Write-Host "Script signed successfully."

# ── 3. Build the EXE with PS2EXE ──────────────────────────────────────────────

if (-not (Get-Module -ListAvailable ps2exe)) {
    Write-Host "Installing ps2exe module ..."
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

Write-Host "Building $OutputExe ..."
Invoke-PS2EXE `
    -InputFile  $SourcePs1 `
    -OutputFile $OutputExe `
    -RequireAdmin `
    -NoConsole:$false

if (-not (Test-Path $OutputExe)) {
    throw "ps2exe did not produce $OutputExe"
}
Write-Host "EXE built."

# ── 4. Sign the EXE with Authenticode ─────────────────────────────────────────

Write-Host "Signing $OutputExe ..."
$sigResult = Set-AuthenticodeSignature `
    -FilePath        $OutputExe `
    -Certificate     $cert `
    -TimestampServer $TimestampServer `
    -HashAlgorithm   SHA256

if ($sigResult.Status -ne "Valid") {
    throw "EXE signing failed: $($sigResult.StatusMessage)"
}

Write-Host ""
Write-Host "Done. Signature details:"
Get-AuthenticodeSignature -FilePath $OutputExe | Format-List Path, Status, StatusMessage, SignerCertificate
