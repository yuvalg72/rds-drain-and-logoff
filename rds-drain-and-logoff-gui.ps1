#Requires -RunAsAdministrator
<#
    Author:   Yuval Grimblat
    Title:    Network Security Solutions Architect and Project Manager
    Company:  Mornex LTD
    Year:     2026
    LinkedIn: https://www.linkedin.com/in/yuvalgrimblat

.SYNOPSIS
    GUI front-end for the RDS Drain and Logoff tool.
    Wraps all drain/logoff operations in a Windows Forms interface with
    live status logging and confirmation dialogs.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Palette ────────────────────────────────────────────────────────────────────
$clrBackground = [System.Drawing.Color]::FromArgb(245, 246, 250)
$clrPanel      = [System.Drawing.Color]::White
$clrAccent     = [System.Drawing.Color]::FromArgb(0, 120, 212)       # Windows blue
$clrDanger     = [System.Drawing.Color]::FromArgb(196, 43, 28)       # red
$clrSuccess    = [System.Drawing.Color]::FromArgb(16, 124, 16)       # green
$clrWarning    = [System.Drawing.Color]::FromArgb(157, 93, 0)        # amber
$clrText       = [System.Drawing.Color]::FromArgb(32, 32, 32)
$clrMuted      = [System.Drawing.Color]::FromArgb(96, 96, 96)
$fontMain      = [System.Drawing.Font]::new("Segoe UI", 9)
$fontBold      = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$fontMono      = [System.Drawing.Font]::new("Consolas", 8.5)
$fontTitle     = [System.Drawing.Font]::new("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)

# ── Helpers ────────────────────────────────────────────────────────────────────

function New-Button {
    param([string]$Text, [System.Drawing.Color]$BackColor, [int]$Width = 160)
    $btn = [System.Windows.Forms.Button]::new()
    $btn.Text      = $Text
    $btn.Width     = $Width
    $btn.Height    = 32
    $btn.FlatStyle = "Flat"
    $btn.BackColor = $BackColor
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.Font      = $fontBold
    $btn.FlatAppearance.BorderSize = 0
    $btn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    return $btn
}

function Write-Log {
    param([string]$Message, [System.Drawing.Color]$Color = $clrText)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logBox.SelectionStart  = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor  = $clrMuted
    $logBox.AppendText("[$timestamp] ")
    $logBox.SelectionColor  = $Color
    $logBox.AppendText("$Message`r`n")
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-ButtonsEnabled([bool]$Enabled) {
    $btnRefresh.Enabled      = $Enabled
    $btnDrainAll.Enabled     = $Enabled
    $btnLogoffAll.Enabled    = $Enabled
    $btnDrainLogoff.Enabled  = $Enabled
}

# ── Data loading ───────────────────────────────────────────────────────────────

function Get-Broker { "$env:COMPUTERNAME.$((Get-CimInstance Win32_ComputerSystem).Domain)" }

function Load-SessionHosts {
    $lvHosts.Items.Clear()
    $broker = $txtBroker.Text.Trim()
    try {
        $collections = Get-RDSessionCollection -ConnectionBroker $broker |
                       Select-Object -ExpandProperty CollectionName
        foreach ($c in $collections) {
            $hosts = Get-RDSessionHost -ConnectionBroker $broker -CollectionName $c |
                     Select-Object SessionHost, NewConnectionAllowed
            foreach ($h in $hosts) {
                $item        = [System.Windows.Forms.ListViewItem]::new($h.SessionHost)
                $statusText  = $h.NewConnectionAllowed
                $subItem     = $item.SubItems.Add($statusText)
                $item.Tag    = $h.SessionHost

                if ($statusText -eq "Yes") {
                    $item.ForeColor = $clrSuccess
                } elseif ($statusText -match "No|Reboot") {
                    $item.ForeColor = $clrDanger
                } else {
                    $item.ForeColor = $clrWarning
                }

                $lvHosts.Items.Add($item) | Out-Null
            }
        }
        Write-Log "Loaded $($lvHosts.Items.Count) session host(s)." $clrSuccess
    } catch {
        Write-Log "Failed to load session hosts: $_" $clrDanger
    }
}

function Load-Sessions {
    $lvSessions.Items.Clear()
    $broker = $txtBroker.Text.Trim()
    try {
        $sessions = Get-RDUserSession -ConnectionBroker $broker
        foreach ($s in $sessions) {
            $item = [System.Windows.Forms.ListViewItem]::new($s.UserName)
            $item.SubItems.Add($s.HostServer)     | Out-Null
            $item.SubItems.Add("$($s.UnifiedSessionId)") | Out-Null
            $item.SubItems.Add($s.SessionState)   | Out-Null
            $item.Tag = $s
            $lvSessions.Items.Add($item) | Out-Null
        }
        $lblSessionCount.Text = "$($lvSessions.Items.Count) active session(s)"
        Write-Log "Loaded $($lvSessions.Items.Count) active session(s)." $clrSuccess
    } catch {
        Write-Log "Failed to load sessions: $_" $clrDanger
    }
}

function Invoke-Refresh {
    Set-ButtonsEnabled $false
    Write-Log "Refreshing data from broker: $($txtBroker.Text.Trim())" $clrAccent
    Load-SessionHosts
    Load-Sessions
    Set-ButtonsEnabled $true
}

# ── Actions ────────────────────────────────────────────────────────────────────

function Invoke-DrainAll {
    if ($lvHosts.Items.Count -eq 0) { Write-Log "No session hosts to drain." $clrWarning; return }
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Set ALL session hosts to drain mode (NotUntilReboot)?`n`nNo new connections will be accepted until each server reboots.",
        "Confirm Drain",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -ne "Yes") { Write-Log "Drain cancelled by user." $clrMuted; return }

    $broker = $txtBroker.Text.Trim()
    Set-ButtonsEnabled $false
    foreach ($item in $lvHosts.Items) {
        $host = $item.Tag
        Write-Log "Draining $host..." $clrAccent
        try {
            Set-RDSessionHost -SessionHost $host `
                              -NewConnectionAllowed NotUntilReboot `
                              -ConnectionBroker $broker
            $item.SubItems[1].Text = "NotUntilReboot"
            $item.ForeColor        = $clrDanger
            Write-Log "Drained: $host" $clrSuccess
        } catch {
            Write-Log "Failed to drain $host`: $_" $clrDanger
        }
    }
    Set-ButtonsEnabled $true
    Write-Log "Drain complete." $clrSuccess
}

function Invoke-LogoffAll {
    if ($lvSessions.Items.Count -eq 0) { Write-Log "No active sessions to log off." $clrWarning; return }
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Force log off ALL $($lvSessions.Items.Count) active session(s)?`n`nUsers will lose unsaved work immediately.",
        "Confirm Logoff",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -ne "Yes") { Write-Log "Logoff cancelled by user." $clrMuted; return }

    $broker = $txtBroker.Text.Trim()
    Set-ButtonsEnabled $false
    foreach ($item in $lvSessions.Items) {
        $s = $item.Tag
        Write-Log "Logging off $($s.UserName) (session $($s.UnifiedSessionId)) on $($s.HostServer)..." $clrAccent
        try {
            Invoke-RDUserLogoff -HostServer       $s.HostServer `
                                -UnifiedSessionID  $s.UnifiedSessionId `
                                -ConnectionBroker  $broker `
                                -Force
            Write-Log "Logged off: $($s.UserName)" $clrSuccess
        } catch {
            Write-Log "Failed to log off $($s.UserName)`: $_" $clrDanger
        }
    }
    Set-ButtonsEnabled $true
    Write-Log "Refreshing session list..." $clrMuted
    Load-Sessions
}

function Invoke-DrainAndLogoff {
    $hostCount    = $lvHosts.Items.Count
    $sessionCount = $lvSessions.Items.Count
    if ($hostCount -eq 0 -and $sessionCount -eq 0) {
        Write-Log "Nothing to drain or log off." $clrWarning; return
    }
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Drain $hostCount host(s) AND force log off $sessionCount session(s)?`n`nThis will immediately terminate all user sessions. Unsaved work will be lost.",
        "Confirm Drain & Logoff",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($result -ne "Yes") { Write-Log "Operation cancelled by user." $clrMuted; return }

    $broker = $txtBroker.Text.Trim()
    Set-ButtonsEnabled $false

    Write-Log "── Draining session hosts ──────────────────" $clrAccent
    foreach ($item in $lvHosts.Items) {
        $host = $item.Tag
        Write-Log "Draining $host..." $clrAccent
        try {
            Set-RDSessionHost -SessionHost $host `
                              -NewConnectionAllowed NotUntilReboot `
                              -ConnectionBroker $broker
            $item.SubItems[1].Text = "NotUntilReboot"
            $item.ForeColor        = $clrDanger
            Write-Log "Drained: $host" $clrSuccess
        } catch {
            Write-Log "Failed to drain $host`: $_" $clrDanger
        }
    }

    Write-Log "── Logging off sessions ────────────────────" $clrAccent
    foreach ($item in $lvSessions.Items) {
        $s = $item.Tag
        Write-Log "Logging off $($s.UserName) (session $($s.UnifiedSessionId))..." $clrAccent
        try {
            Invoke-RDUserLogoff -HostServer       $s.HostServer `
                                -UnifiedSessionID  $s.UnifiedSessionId `
                                -ConnectionBroker  $broker `
                                -Force
            Write-Log "Logged off: $($s.UserName)" $clrSuccess
        } catch {
            Write-Log "Failed to log off $($s.UserName)`: $_" $clrDanger
        }
    }

    Set-ButtonsEnabled $true
    Write-Log "Refreshing session list..." $clrMuted
    Load-Sessions
    Write-Log "Drain and logoff complete." $clrSuccess
}

# ── Form ───────────────────────────────────────────────────────────────────────

$form                  = [System.Windows.Forms.Form]::new()
$form.Text             = "RDS Drain and Logoff Tool  –  Mornex LTD"
$form.Size             = [System.Drawing.Size]::new(900, 680)
$form.MinimumSize      = [System.Drawing.Size]::new(800, 600)
$form.StartPosition    = "CenterScreen"
$form.BackColor        = $clrBackground
$form.Font             = $fontMain
$form.Icon             = [System.Drawing.SystemIcons]::Shield

# ── Header ─────────────────────────────────────────────────────────────────────

$pnlHeader             = [System.Windows.Forms.Panel]::new()
$pnlHeader.Dock        = "Top"
$pnlHeader.Height      = 56
$pnlHeader.BackColor   = $clrAccent
$pnlHeader.Padding     = [System.Windows.Forms.Padding]::new(16, 0, 16, 0)

$lblTitle              = [System.Windows.Forms.Label]::new()
$lblTitle.Text         = "RDS Drain and Logoff Tool"
$lblTitle.Font         = $fontTitle
$lblTitle.ForeColor    = [System.Drawing.Color]::White
$lblTitle.AutoSize     = $true
$lblTitle.Location     = [System.Drawing.Point]::new(16, 10)

$lblAuthor             = [System.Windows.Forms.Label]::new()
$lblAuthor.Text        = "Yuval Grimblat  |  Network Security Solutions Architect & PM  |  Mornex LTD 2026"
$lblAuthor.Font        = [System.Drawing.Font]::new("Segoe UI", 8)
$lblAuthor.ForeColor   = [System.Drawing.Color]::FromArgb(200, 230, 255)
$lblAuthor.AutoSize    = $true
$lblAuthor.Location    = [System.Drawing.Point]::new(16, 34)

$pnlHeader.Controls.AddRange(@($lblTitle, $lblAuthor))

# ── Broker bar ─────────────────────────────────────────────────────────────────

$pnlBroker             = [System.Windows.Forms.Panel]::new()
$pnlBroker.Dock        = "Top"
$pnlBroker.Height      = 48
$pnlBroker.BackColor   = $clrPanel
$pnlBroker.Padding     = [System.Windows.Forms.Padding]::new(12, 8, 12, 8)

$lblBrokerLabel        = [System.Windows.Forms.Label]::new()
$lblBrokerLabel.Text   = "Connection Broker:"
$lblBrokerLabel.Font   = $fontBold
$lblBrokerLabel.AutoSize = $true
$lblBrokerLabel.Location = [System.Drawing.Point]::new(12, 14)

$txtBroker             = [System.Windows.Forms.TextBox]::new()
$txtBroker.Location    = [System.Drawing.Point]::new(148, 11)
$txtBroker.Width       = 360
$txtBroker.Height      = 24
$txtBroker.BorderStyle = "FixedSingle"
$txtBroker.Font        = $fontMono
$txtBroker.Text        = (Get-Broker)

$btnRefresh            = New-Button "  Refresh" $clrAccent 110
$btnRefresh.Location   = [System.Drawing.Point]::new(520, 8)
$btnRefresh.Image      = [System.Drawing.SystemIcons]::Information.ToBitmap()
$btnRefresh.ImageAlign = "MiddleLeft"
$btnRefresh.TextAlign  = "MiddleRight"
$btnRefresh.Image      = $null  # keep text-only for simplicity

$pnlBroker.Controls.AddRange(@($lblBrokerLabel, $txtBroker, $btnRefresh))

# ── Main split ─────────────────────────────────────────────────────────────────

$split                 = [System.Windows.Forms.SplitContainer]::new()
$split.Dock            = "Fill"
$split.Orientation     = "Vertical"
$split.SplitterWidth   = 6
$split.SplitterDistance = 340
$split.BackColor       = $clrBackground
$split.Panel1.Padding  = [System.Windows.Forms.Padding]::new(12, 8, 6, 8)
$split.Panel2.Padding  = [System.Windows.Forms.Padding]::new(6, 8, 12, 8)

# Session Hosts panel
$pnlHosts              = [System.Windows.Forms.GroupBox]::new()
$pnlHosts.Text         = "Session Hosts"
$pnlHosts.Dock         = "Fill"
$pnlHosts.Font         = $fontBold
$pnlHosts.ForeColor    = $clrText
$pnlHosts.BackColor    = $clrPanel

$lvHosts               = [System.Windows.Forms.ListView]::new()
$lvHosts.Dock          = "Fill"
$lvHosts.View          = "Details"
$lvHosts.FullRowSelect = $true
$lvHosts.GridLines     = $true
$lvHosts.Font          = $fontMain
$lvHosts.BorderStyle   = "None"
$lvHosts.BackColor     = $clrPanel
$lvHosts.Columns.Add("Host", 200)    | Out-Null
$lvHosts.Columns.Add("Status", 120)  | Out-Null

$pnlHosts.Controls.Add($lvHosts)
$split.Panel1.Controls.Add($pnlHosts)

# Active Sessions panel
$pnlSessions           = [System.Windows.Forms.GroupBox]::new()
$pnlSessions.Text      = "Active Sessions"
$pnlSessions.Dock      = "Fill"
$pnlSessions.Font      = $fontBold
$pnlSessions.ForeColor = $clrText
$pnlSessions.BackColor = $clrPanel

$lblSessionCount       = [System.Windows.Forms.Label]::new()
$lblSessionCount.Text  = "0 active session(s)"
$lblSessionCount.Font  = [System.Drawing.Font]::new("Segoe UI", 8)
$lblSessionCount.ForeColor = $clrMuted
$lblSessionCount.AutoSize  = $true
$lblSessionCount.Anchor    = "Top,Right"

$lvSessions            = [System.Windows.Forms.ListView]::new()
$lvSessions.Dock       = "Fill"
$lvSessions.View       = "Details"
$lvSessions.FullRowSelect = $true
$lvSessions.GridLines  = $true
$lvSessions.Font       = $fontMain
$lvSessions.BorderStyle = "None"
$lvSessions.BackColor  = $clrPanel
$lvSessions.Columns.Add("User", 130)       | Out-Null
$lvSessions.Columns.Add("Host", 130)       | Out-Null
$lvSessions.Columns.Add("Session ID", 80)  | Out-Null
$lvSessions.Columns.Add("State", 80)       | Out-Null

$pnlSessions.Controls.Add($lvSessions)
$split.Panel2.Controls.Add($pnlSessions)

# ── Action buttons ─────────────────────────────────────────────────────────────

$pnlActions            = [System.Windows.Forms.Panel]::new()
$pnlActions.Dock       = "Bottom"
$pnlActions.Height     = 52
$pnlActions.BackColor  = $clrPanel
$pnlActions.Padding    = [System.Windows.Forms.Padding]::new(12, 10, 12, 10)

$btnDrainAll           = New-Button "Drain All Hosts"       $clrWarning 160
$btnDrainAll.Location  = [System.Drawing.Point]::new(12, 10)

$btnLogoffAll          = New-Button "Logoff All Sessions"   $clrDanger  170
$btnLogoffAll.Location = [System.Drawing.Point]::new(184, 10)

$btnDrainLogoff        = New-Button "Drain + Logoff All"    $clrDanger  170
$btnDrainLogoff.Location = [System.Drawing.Point]::new(366, 10)

$pnlActions.Controls.AddRange(@($btnDrainAll, $btnLogoffAll, $btnDrainLogoff))

# ── Status log ─────────────────────────────────────────────────────────────────

$pnlLog                = [System.Windows.Forms.GroupBox]::new()
$pnlLog.Text           = "Status Log"
$pnlLog.Dock           = "Bottom"
$pnlLog.Height         = 180
$pnlLog.Font           = $fontBold
$pnlLog.ForeColor      = $clrText
$pnlLog.BackColor      = $clrPanel
$pnlLog.Padding        = [System.Windows.Forms.Padding]::new(4)

$logBox                = [System.Windows.Forms.RichTextBox]::new()
$logBox.Dock           = "Fill"
$logBox.ReadOnly       = $true
$logBox.Font           = $fontMono
$logBox.BackColor      = [System.Drawing.Color]::FromArgb(30, 30, 30)
$logBox.ForeColor      = [System.Drawing.Color]::FromArgb(220, 220, 220)
$logBox.BorderStyle    = "None"
$logBox.ScrollBars     = "Vertical"
$pnlLog.Controls.Add($logBox)

# ── LinkedIn link ───────────────────────────────────────────────────────────────

$lnkLinkedIn           = [System.Windows.Forms.LinkLabel]::new()
$lnkLinkedIn.Text      = "linkedin.com/in/yuvalgrimblat"
$lnkLinkedIn.Dock      = "Bottom"
$lnkLinkedIn.Height    = 20
$lnkLinkedIn.TextAlign = "MiddleRight"
$lnkLinkedIn.Font      = [System.Drawing.Font]::new("Segoe UI", 8)
$lnkLinkedIn.LinkColor = $clrAccent
$lnkLinkedIn.Padding   = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
$lnkLinkedIn.Add_LinkClicked({
    Start-Process "https://www.linkedin.com/in/yuvalgrimblat"
})

# ── Assemble form ───────────────────────────────────────────────────────────────

$form.Controls.Add($split)
$form.Controls.Add($pnlLog)
$form.Controls.Add($pnlActions)
$form.Controls.Add($lnkLinkedIn)
$form.Controls.Add($pnlBroker)
$form.Controls.Add($pnlHeader)

# ── Wire events ────────────────────────────────────────────────────────────────

$btnRefresh.Add_Click(    { Invoke-Refresh })
$btnDrainAll.Add_Click(   { Invoke-DrainAll })
$btnLogoffAll.Add_Click(  { Invoke-LogoffAll })
$btnDrainLogoff.Add_Click({ Invoke-DrainAndLogoff })

$form.Add_Shown({
    Write-Log "RDS Drain and Logoff Tool started." $clrAccent
    Write-Log "Author: Yuval Grimblat | Mornex LTD 2026" $clrMuted
    Write-Log "Connecting to broker: $($txtBroker.Text.Trim())" $clrAccent
    Invoke-Refresh
})

# ── Run ────────────────────────────────────────────────────────────────────────

[System.Windows.Forms.Application]::Run($form)
