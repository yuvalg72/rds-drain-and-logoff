#Requires -RunAsAdministrator
<#
    Author:   Yuval Grimblat
    Title:    Network Security Solutions Architect and Project Manager
    Company:  Mornex LTD
    Year:     2026
    LinkedIn: https://www.linkedin.com/in/yuvalgrimblat

.SYNOPSIS
    Builds rds-drain-and-logoff.exe from the PowerShell source and signs it with Authenticode.

.DESCRIPTION
    Supports three certificate sources:
      1. SelfSigned  – create a self-signed certificate (default; must be trusted via
                       Group Policy or manually on each target machine)
      2. InternalCA  – request a certificate from your enterprise (ADCS) CA
      3. Thumbprint  – use an existing certificate already in your cert store

.PARAMETER CertSource
    "SelfSigned" | "InternalCA" | "Thumbprint"   (default: SelfSigned)

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
    # Quick self-signed build (default, testing / internal lab)
    .\build-and-sign.ps1

.EXAMPLE
    # Issue a cert from your internal Active Directory Certificate Services
    .\build-and-sign.ps1 -CertSource InternalCA -CATemplate "CodeSigning"

.EXAMPLE
    # Re-use an existing cert already in your store
    .\build-and-sign.ps1 -CertSource Thumbprint -Thumbprint "AABBCC..."
#>

param(
    [ValidateSet("SelfSigned", "InternalCA", "Thumbprint")]
    [string] $CertSource      = "SelfSigned",

    [string] $Thumbprint      = "",
    [string] $CATemplate      = "CodeSigning",
    [string] $CertSubject     = "RDS Drain and Logoff Tool",
    [string] $TimestampServer = "http://timestamp.digicert.com"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# $PSScriptRoot is empty when the script is dot-sourced; fall back to the current directory.
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$SourcePs1 = Join-Path $ScriptDir "drain-rds-farm.ps1"
$OutputExe = Join-Path $ScriptDir "rds-drain-and-logoff.exe"

# ── 1. Resolve the signing certificate ────────────────────────────────────────

function Get-SigningCert([string]$Source, [string]$Print, [string]$Template, [string]$Subject) {
    switch ($Source) {

        "Thumbprint" {
            if (-not $Print) {
                throw "Provide -Thumbprint when using CertSource=Thumbprint."
            }
            $cert = (Get-ChildItem Cert:\CurrentUser\My, Cert:\LocalMachine\My -ErrorAction SilentlyContinue) |
                    Where-Object { $_.Thumbprint -eq $Print } |
                    Select-Object -First 1
            if (-not $cert) {
                throw "Certificate with thumbprint '$Print' not found in any local store."
            }
            return $cert
        }

        "InternalCA" {
            Write-Host "Requesting certificate from internal CA (template: $Template)..."
            $request = Get-Certificate `
                -Template          $Template `
                -SubjectName       "CN=$Subject" `
                -CertStoreLocation Cert:\CurrentUser\My
            return $request.Certificate
        }

        "SelfSigned" {
            Write-Warning "Self-signed certificate created. Add it to 'Trusted Publishers' and 'Trusted Root CA' stores on every machine that will run the EXE."
            $cert = New-SelfSignedCertificate `
                -Type              CodeSigningCert `
                -Subject           "CN=$Subject" `
                -HashAlgorithm     SHA256 `
                -KeyExportPolicy   Exportable `
                -CertStoreLocation Cert:\CurrentUser\My `
                -NotAfter          (Get-Date).AddYears(3)

            # Trust locally so Defender accepts the signed binary on this machine immediately.
            foreach ($storeName in @("Root", "TrustedPublisher")) {
                $store = [System.Security.Cryptography.X509Certificates.X509Store]::new(
                    $storeName, "LocalMachine")
                try {
                    $store.Open("ReadWrite")
                    $store.Add($cert)
                } finally {
                    $store.Close()
                }
            }

            Write-Host "Self-signed cert added to local Trusted Root CA and Trusted Publishers stores."
            return $cert
        }

        default {
            throw "Unknown CertSource '$Source'."
        }
    }
}

$cert = Get-SigningCert -Source $CertSource -Print $Thumbprint -Template $CATemplate -Subject $CertSubject
Write-Host "Using certificate: Subject=$($cert.Subject)  Thumbprint=$($cert.Thumbprint)"

# ── 2. Build the EXE with PS2EXE ──────────────────────────────────────────────
# Build BEFORE signing the source script so PS2EXE does not embed the
# Authenticode signature block as script text inside the EXE.

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing ps2exe module ..."
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

Write-Host "Building $OutputExe ..."
Invoke-PS2EXE `
    -InputFile  $SourcePs1 `
    -OutputFile $OutputExe `
    -RequireAdmin

if (-not (Test-Path $OutputExe)) {
    throw "ps2exe did not produce $OutputExe"
}
Write-Host "EXE built."

# ── 3. Sign the PowerShell source script and the EXE ──────────────────────────

foreach ($target in @($SourcePs1, $OutputExe)) {
    Write-Host "Signing $target ..."
    $sigResult = Set-AuthenticodeSignature `
        -FilePath        $target `
        -Certificate     $cert `
        -TimestampServer $TimestampServer `
        -HashAlgorithm   SHA256

    if ($sigResult.Status -ne "Valid") {
        throw "Signing failed for $target`: $($sigResult.StatusMessage)"
    }
    Write-Host "Signed: $target"
}

Write-Host ""
Write-Host "Done. Signature details:"
Get-AuthenticodeSignature -FilePath $OutputExe | Format-List Path, Status, StatusMessage, SignerCertificate
