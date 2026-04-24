# RDS Drain and Logoff

PowerShell utility for placing all **Remote Desktop Services (RDS) Session Hosts** into **drain mode** and forcing logoff of active user sessions across all collections through the **RD Connection Broker**.

This tool is designed for **maintenance operations** such as patching, updates, infrastructure changes, or scheduled downtime where all active sessions must be terminated and new connections blocked.

---

# Overview

The script performs the following actions:

1. Detects the **Active Directory domain** of the host running the script.
2. Identifies the **RD Connection Broker**.
3. Retrieves all **RDS Session Collections**.
4. Enumerates all **Session Hosts** across those collections.
5. Sets each Session Host to **Drain Mode** (`NotUntilReboot`).
6. Enumerates all active **RDS user sessions**.
7. Forces **logoff of all connected users**.

---

# Architecture Flow

```
Connection Broker
        │
        │
        ▼
Get-RDSessionCollection
        │
        ▼
Get-RDSessionHost
        │
        ▼
Drain Session Hosts
(NewConnectionAllowed = NotUntilReboot)
        │
        ▼
Enumerate Active Sessions
        │
        ▼
Force Logoff
```

---

# Use Cases

Typical scenarios where this tool is used:

* RDS maintenance windows
* Windows Updates on RDS servers
* Infrastructure upgrades
* Storage migrations
* Security patching
* Emergency shutdown of user sessions
* Pre-reboot draining of session hosts

---

# Script Behavior

The script performs two major actions:

## 1. Drain Session Hosts

Each RDS Session Host is configured with:

```
NewConnectionAllowed = NotUntilReboot
```

Meaning:

* No new user connections are allowed
* The server remains in this state until reboot

This ensures that:

* Users cannot reconnect to the host
* Maintenance can proceed safely

---

## 2. Forced User Logoff

The script enumerates all active RDS sessions and executes:

```
Invoke-RDUserLogoff -Force
```

This will immediately terminate all active user sessions.

⚠ Users **will not receive a warning**, and unsaved work may be lost.

---

# Requirements

The following prerequisites must be met:

* Windows Server with **Remote Desktop Services**
* PowerShell with **RemoteDesktop module**
* Administrative privileges
* Access to the **RD Connection Broker**
* Execution within the RDS infrastructure domain

Required PowerShell commands used:

```
Get-RDSessionCollection
Get-RDSessionHost
Set-RDSessionHost
Get-RDUserSession
Invoke-RDUserLogoff
```

---

# Usage

Run the script from a server that can communicate with the **RD Connection Broker**.

```
powershell.exe -ExecutionPolicy Bypass -File rds-drain-and-logoff.ps1
```

---

# Expected Output

Example console output:

```
Configuring RDSH01...
Completed configuration for RDSH01

Configuring RDSH02...
Completed configuration for RDSH02

Logging off user sessions...
User session terminated: SessionID 4
User session terminated: SessionID 7
```

---

# Safety Warning

This script is **destructive to active sessions**.

It will:

* Immediately log off all users
* Terminate running applications
* Potentially cause loss of unsaved data

Recommended precautions:

* Run during maintenance windows
* Notify users before execution
* Verify backups and recovery procedures

---

# Building & Signing the EXE

The script can be packaged as a signed executable using **PS2EXE** and
**Authenticode** (Windows built-in code signing).  A valid Authenticode
signature is required by Microsoft Defender and most EDR solutions before they
will allow an unsigned, locally-built binary to run.

Use the included `build-and-sign.ps1` helper script.  Run it from an elevated
(Administrator) PowerShell prompt.

---

## Option 1 – Self-signed certificate (default)

No certificate infrastructure needed.  Just run the script with no arguments:

```powershell
.\build-and-sign.ps1
```

This creates a self-signed code-signing certificate, trusts it locally, builds
the EXE, and signs it in one step.  Defender will accept the EXE **on the
machine where you ran the script**.

To allow the EXE to run on other machines, distribute the certificate via Group
Policy:

```
Computer Configuration → Windows Settings → Security Settings →
Public Key Policies → Trusted Publishers   (add the .cer export)
                     → Trusted Root CAs    (add the .cer export)
```

Export the certificate for distribution:

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -match "RDS Drain" }
Export-Certificate -Cert $cert -FilePath rds-drain-logoff-signing.cer -Type CERT
```

---

## Option 2 – Internal Active Directory Certificate Services (ADCS)

If your domain has an ADCS CA with a CodeSigning template, the script will
automatically request and retrieve a certificate:

```powershell
.\build-and-sign.ps1 -CertSource InternalCA -CATemplate "CodeSigning"
```

The resulting certificate is trusted by every machine that trusts your domain
CA — no extra trust distribution is needed.

---

## Option 3 – Existing certificate

If your organisation already issued you a code-signing certificate, pass its
thumbprint:

```powershell
.\build-and-sign.ps1 -CertSource Thumbprint -Thumbprint "AABBCCDDEEFF..."
```

To find the thumbprint of certificates already in your store:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.EnhancedKeyUsageList -match "Code Signing" }
```

---

## Verifying the signature

```powershell
Get-AuthenticodeSignature .\rds-drain-and-logoff.exe | Format-List
```

A correctly signed binary shows `Status: Valid`.

---

# Recommended Improvements

Future versions may include:

* User warning notifications
* Graceful session draining
* Maintenance countdown timer
* Logging to file
* Error handling
* Session filtering
* Dry-run mode
* Integration with monitoring systems

---

# Version

```
Version: 1.0
Year:     2026
Author:   Yuval Grimblat
Title:    Network Security Solutions Architect and Project Manager
Company:  Mornex LTD
LinkedIn: https://www.linkedin.com/in/yuvalgrimblat
```

---

# License

MIT License

---
