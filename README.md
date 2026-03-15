להלן **README מקצועי ועדכני (2026)** שמתאים לריפו בשם
`rds-drain-and-logoff`

ניתן להעתיק ישירות ל-`README.md`.

---

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

# Optional: Build EXE

The script can be packaged as an executable using **PS2EXE**.

Install the module:

```
Install-Module ps2exe
```

Build executable:

```
Invoke-PS2EXE rds-drain-and-logoff.ps1 rds-drain-and-logoff.exe -RequireAdmin
```

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
Year: 2026
Author: Yuval Grimblat
```

---

# License

MIT License

---
