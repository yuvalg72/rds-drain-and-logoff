# RDS Drain and Logoff

A PowerShell toolkit for placing **Remote Desktop Services (RDS) Session Hosts** into drain mode and force-logging off all active user sessions across every collection managed by an **RD Connection Broker**.

Built for maintenance operations: patching, infrastructure upgrades, storage migrations, and emergency session termination.

---

## Files

| File | Purpose |
|---|---|
| `rds-drain-and-logoff-gui.ps1` | **GUI front-end** — 2026 dark-mode Windows Forms interface (recommended) |
| `drain-rds-farm.ps1` | **CLI script** — headless, runs from any elevated PowerShell prompt |
| `build-and-sign.ps1` | **Build & sign helper** — packages the CLI into a signed EXE via PS2EXE |

---

## Requirements

* Windows Server with **Remote Desktop Services** role installed
* PowerShell **RemoteDesktop** module (`RSAT-RDS-Tools` feature)
* Run from a machine with network access to the **RD Connection Broker**
* **Administrator** privileges (`#Requires -RunAsAdministrator` enforced on all scripts)
* GUI additionally requires **.NET Windows Forms** (included in all Windows Server editions)

---

## GUI — Recommended

Launch from an elevated PowerShell prompt:

```powershell
powershell.exe -ExecutionPolicy Bypass -File rds-drain-and-logoff-gui.ps1
```

### Interface

```
┌─ Purple → Blue gradient header ──────────────────────────────────┐
│  RDS Drain and Logoff   Yuval Grimblat · Mornex LTD 2026  [LinkedIn]
├─ Connection Broker: [BROKER.DOMAIN.COM]  [⟳ Refresh] ───────────┤
├─ Session Hosts  [3] ──────────┬─ Active Sessions  [7] ───────────┤
│  RDSH01  NotUntilReboot (red) │  jsmith   RDSH01  4   Active     │
│  RDSH02  Yes           (green)│  bjones   RDSH02  7   Active     │
│  RDSH03  NotUntilReboot (red) │  ...                             │
├───────────────────────────────┴──────────────────────────────────┤
│  [⬇ Drain All Hosts]  [✕ Logoff All Sessions]  [⚡ Drain+Logoff] │
├─ Status Log ─────────────────────────────────────────────────────┤
│  14:02:01  Connecting to broker: BROKER.DOMAIN.COM…              │
│  14:02:02  Loaded 3 session host(s).                             │
│  14:02:02  Loaded 7 active session(s).                           │
└──────────────────────────────────────────────────────────────────┘
```

### Features

**Connection Broker**
- Auto-detected from the local machine's domain on startup
- Editable — point it at any broker in your environment
- Refresh reloads both lists from the broker on demand

**Session Hosts panel**
- Colour-coded status: green = accepting connections, red = drained (`NotUntilReboot`), amber = other
- Live badge showing host count
- Updates in place after drain operations

**Active Sessions panel**
- Shows username, host server, session ID, and session state
- Live badge showing session count
- Reloads automatically after any logoff operation

**Action buttons** (each requires confirmation before executing)

| Button | Action |
|---|---|
| Drain All Hosts | Sets every session host to `NotUntilReboot` |
| Logoff All Sessions | Force-logs off every listed session |
| Drain + Logoff All | Drains all hosts then logs off all sessions in sequence |

**Status Log**
- Timestamped, colour-coded output for every operation
- Purple section dividers, blue for in-progress, green for success, red for errors
- Dark terminal style (Cascadia Code / Consolas)

**Design**
- 2026 dark-mode aesthetic: near-black background, rounded card panels, purple-to-blue gradient header, pill-shaped buttons with hover transitions
- Fully responsive — all panels reflow on window resize
- Clickable LinkedIn link in the header

---

## CLI — Headless

Run from any elevated PowerShell session on a machine with broker access:

```powershell
powershell.exe -ExecutionPolicy Bypass -File drain-rds-farm.ps1
```

### What it does

1. Detects the Active Directory domain via `Get-CimInstance Win32_ComputerSystem`
2. Constructs the Connection Broker FQDN (`$env:COMPUTERNAME.$domain`)
3. Retrieves all RDS Session Collections from the broker
4. Enumerates all Session Hosts across those collections
5. Sets every Session Host to `NewConnectionAllowed = NotUntilReboot`
6. Enumerates all active RDS user sessions via the broker
7. Force-logs off every session with `Invoke-RDUserLogoff -Force`

Errors on individual hosts or sessions are caught and reported without aborting the remaining operations.

### Example output

```
Configuring RDSH01...
Completed configuration for RDSH01
Configuring RDSH02...
Completed configuration for RDSH02
Logged off session 4 on RDSH01
Logged off session 7 on RDSH02
```

---

## Architecture Flow

```
                    ┌─────────────────────┐
                    │   RD Connection     │
                    │      Broker         │
                    └──────────┬──────────┘
                               │
               ┌───────────────┼───────────────┐
               ▼               ▼               ▼
   Get-RDSessionCollection  (all collections enumerated)
               │
               ▼
       Get-RDSessionHost
               │
               ▼
     Set-RDSessionHost
   NewConnectionAllowed
    = NotUntilReboot
               │
               ▼
      Get-RDUserSession
               │
               ▼
   Invoke-RDUserLogoff -Force
```

---

## Building a Signed EXE

The CLI script can be compiled into a signed Windows executable using **PS2EXE** and **Authenticode**.  A valid Authenticode signature is required by Microsoft Defender and most EDR solutions before they will allow a locally-built binary to run.

Use `build-and-sign.ps1` from an elevated PowerShell prompt.  It builds the EXE first and then signs both the `.ps1` source and the `.exe` output so no signature text is embedded into the compiled binary.

### Option 1 — Self-signed (default)

No certificate infrastructure needed:

```powershell
.\build-and-sign.ps1
```

Creates a self-signed code-signing certificate, trusts it on the local machine, builds the EXE, and signs both files. Defender accepts the EXE immediately on the machine where the script ran.

To run the EXE on other machines, export the certificate and distribute it via Group Policy:

```powershell
# Export
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -match "RDS Drain" }
Export-Certificate -Cert $cert -FilePath rds-drain-logoff-signing.cer -Type CERT
```

```
GPO path: Computer Configuration → Windows Settings → Security Settings →
          Public Key Policies → Trusted Publishers  (import .cer)
                               → Trusted Root CAs   (import .cer)
```

### Option 2 — Internal ADCS CA

Automatically requests a certificate from your domain's Active Directory Certificate Services:

```powershell
.\build-and-sign.ps1 -CertSource InternalCA -CATemplate "CodeSigning"
```

The resulting certificate is trusted by every domain-joined machine — no manual distribution needed.

### Option 3 — Existing certificate

```powershell
.\build-and-sign.ps1 -CertSource Thumbprint -Thumbprint "AABBCCDDEEFF..."
```

Find available code-signing certificates:

```powershell
Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.EnhancedKeyUsageList -match "Code Signing" }
```

### Verify the signature

```powershell
Get-AuthenticodeSignature .\rds-drain-and-logoff.exe | Format-List
# Status should show: Valid
```

---

## Safety Warning

This tool is **destructive to active sessions**.

- Users are logged off **immediately** with no prior warning
- Running applications are terminated
- Unsaved work **will be lost**

Recommended precautions:
- Run during a scheduled maintenance window
- Notify users before execution
- Use the GUI confirmation dialogs as a last checkpoint

---

## Version

```
Version:  1.1
Year:     2026
Author:   Yuval Grimblat
Title:    Network Security Solutions Architect and Project Manager
Company:  Mornex LTD
LinkedIn: https://www.linkedin.com/in/yuvalgrimblat
```

---

## License

MIT License
