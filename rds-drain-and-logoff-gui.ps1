#Requires -RunAsAdministrator
<#
    Version:  1.1
    Author:   Yuval Grimblat
    Title:    Network Security Solutions Architect and Project Manager
    Company:  Mornex LTD
    Year:     2026
    LinkedIn: https://www.linkedin.com/in/yuvalgrimblat
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Custom controls (rounded panels + pill buttons) ───────────────────────────
Add-Type @"
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class RoundedPanel : Panel {
    public int CornerRadius { get; set; } = 14;
    public Color BorderColor { get; set; } = Color.FromArgb(48, 54, 61);
    public int BorderWidth  { get; set; } = 1;

    private GraphicsPath BuildPath(Rectangle r, int rad) {
        var p = new GraphicsPath();
        p.AddArc(r.X,               r.Y,                rad*2, rad*2, 180, 90);
        p.AddArc(r.Right - rad*2,   r.Y,                rad*2, rad*2, 270, 90);
        p.AddArc(r.Right - rad*2,   r.Bottom - rad*2,   rad*2, rad*2,   0, 90);
        p.AddArc(r.X,               r.Bottom - rad*2,   rad*2, rad*2,  90, 90);
        p.CloseFigure();
        return p;
    }

    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(1, 1, Width - 2, Height - 2);
        using (var path   = BuildPath(rect, CornerRadius))
        using (var fill   = new SolidBrush(BackColor))
        using (var border = new Pen(BorderColor, BorderWidth)) {
            Region = new Region(path);
            e.Graphics.FillPath(fill,   path);
            e.Graphics.DrawPath(border, path);
        }
        base.OnPaint(e);
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); Invalidate(); }
}

public class PillButton : Button {
    public Color HoverColor  { get; set; }
    private bool _hovered = false;

    public PillButton() {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint
               | ControlStyles.DoubleBuffer, true);
        Cursor = Cursors.Hand;
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
    }

    protected override void OnMouseEnter(EventArgs e) { _hovered = true;  Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e) {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(0, 0, Width - 1, Height - 1);
        int r = Height / 2;
        var path = new GraphicsPath();
        path.AddArc(rect.X,             rect.Y,             r*2, r*2, 180, 90);
        path.AddArc(rect.Right - r*2,   rect.Y,             r*2, r*2, 270, 90);
        path.AddArc(rect.Right - r*2,   rect.Bottom - r*2,  r*2, r*2,   0, 90);
        path.AddArc(rect.X,             rect.Bottom - r*2,  r*2, r*2,  90, 90);
        path.CloseFigure();

        Region = new Region(path);
        Color fill = _hovered && HoverColor != Color.Empty ? HoverColor : BackColor;
        using (var brush = new SolidBrush(fill))
            e.Graphics.FillPath(brush, path);

        using (var sf  = new StringFormat { Alignment = StringAlignment.Center,
                                            LineAlignment = StringAlignment.Center })
        using (var tb  = new SolidBrush(ForeColor))
            e.Graphics.DrawString(Text, Font, tb, new RectangleF(0,0,Width,Height), sf);

        path.Dispose();
    }
}

public class DarkListView : ListView {
    public DarkListView() {
        SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
        OwnerDraw    = true;
        View         = View.Details;
        FullRowSelect = true;
        GridLines    = false;
        BorderStyle  = BorderStyle.None;
    }

    protected override void OnDrawColumnHeader(DrawListViewColumnHeaderEventArgs e) {
        using (var bg = new SolidBrush(Color.FromArgb(22, 27, 34)))
        using (var fg = new SolidBrush(Color.FromArgb(132, 141, 151)))
        using (var sf = new StringFormat { LineAlignment = StringAlignment.Center,
                                           Alignment = StringAlignment.Near,
                                           FormatFlags = StringFormatFlags.NoWrap }) {
            e.Graphics.FillRectangle(bg, e.Bounds);
            var textRect = new RectangleF(e.Bounds.X + 8, e.Bounds.Y,
                                          e.Bounds.Width - 8, e.Bounds.Height);
            e.Graphics.DrawString(e.Header.Text, e.Font, fg, textRect, sf);
        }
    }

    protected override void OnDrawItem(DrawListViewItemEventArgs e) { e.DrawDefault = true; }

    protected override void OnDrawSubItem(DrawListViewSubItemEventArgs e) {
        bool selected = e.Item.Selected;
        Color bg = selected ? Color.FromArgb(33, 81, 133)
                            : (e.ItemIndex % 2 == 0 ? Color.FromArgb(22, 27, 34)
                                                     : Color.FromArgb(26, 32, 41));
        using (var brush = new SolidBrush(bg))
            e.Graphics.FillRectangle(brush, e.Bounds);

        using (var fg = new SolidBrush(e.Item.ForeColor))
        using (var sf = new StringFormat { LineAlignment = StringAlignment.Center,
                                           FormatFlags   = StringFormatFlags.NoWrap,
                                           Trimming      = StringTrimming.EllipsisCharacter }) {
            var textRect = new RectangleF(e.Bounds.X + 6, e.Bounds.Y,
                                          e.Bounds.Width - 6, e.Bounds.Height);
            e.Graphics.DrawString(e.SubItem.Text, e.Item.Font, fg, textRect, sf);
        }
    }
}
"@ -ReferencedAssemblies System.Windows.Forms, System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# ── Palette ────────────────────────────────────────────────────────────────────
$BG        = [System.Drawing.Color]::FromArgb(13,  17,  23)
$SURFACE   = [System.Drawing.Color]::FromArgb(22,  27,  34)
$CARD      = [System.Drawing.Color]::FromArgb(33,  38,  45)
$BORDER    = [System.Drawing.Color]::FromArgb(48,  54,  61)
$ACCENT    = [System.Drawing.Color]::FromArgb(47, 129, 247)
$PURPLE    = [System.Drawing.Color]::FromArgb(139, 92, 246)
$SUCCESS   = [System.Drawing.Color]::FromArgb(63, 185,  80)
$WARNING   = [System.Drawing.Color]::FromArgb(210, 153,  34)
$DANGER    = [System.Drawing.Color]::FromArgb(248,  81,  73)
$TEXT      = [System.Drawing.Color]::FromArgb(230, 237, 243)
$MUTED     = [System.Drawing.Color]::FromArgb(132, 141, 151)

$fUI       = [System.Drawing.Font]::new("Segoe UI",  9)
$fBold     = [System.Drawing.Font]::new("Segoe UI",  9,  [System.Drawing.FontStyle]::Bold)
$fTitle    = [System.Drawing.Font]::new("Segoe UI", 13,  [System.Drawing.FontStyle]::Bold)
$fSub      = [System.Drawing.Font]::new("Segoe UI",  8)
$fMono     = [System.Drawing.Font]::new("Cascadia Code", 8.5)
if (-not (Test-Path "C:\Windows\Fonts\CascadiaCode.ttf")) {
    $fMono = [System.Drawing.Font]::new("Consolas", 9)
}

# ── Helpers ────────────────────────────────────────────────────────────────────

function New-PillBtn([string]$Text, [System.Drawing.Color]$Color,
                     [System.Drawing.Color]$Hover, [int]$W = 170) {
    $b = [PillButton]::new()
    $b.Text       = $Text
    $b.Width      = $W
    $b.Height     = 36
    $b.BackColor  = $Color
    $b.HoverColor = $Hover
    $b.ForeColor  = $TEXT
    $b.Font       = $fBold
    return $b
}

function New-Card([int]$Radius = 14) {
    $p = [RoundedPanel]::new()
    $p.BackColor    = $CARD
    $p.BorderColor  = $BORDER
    $p.CornerRadius = $Radius
    return $p
}

function New-Label([string]$Text, $Font = $null, $Color = $null) {
    $l = [System.Windows.Forms.Label]::new()
    $l.Text      = $Text
    $l.AutoSize  = $true
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.ForeColor = if ($Color) { $Color } else { $TEXT }
    $l.Font      = if ($Font)  { $Font  } else { $fUI  }
    return $l
}

function Write-Log([string]$Msg, $Color = $null) {
    if (-not $Color) { $Color = $TEXT }
    $ts = Get-Date -Format "HH:mm:ss"
    $logBox.SelectionStart  = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor  = $MUTED
    $logBox.AppendText("$ts  ")
    $logBox.SelectionColor  = $Color
    $logBox.AppendText("$Msg`r`n")
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-UI([bool]$On) {
    foreach ($b in @($btnRefresh, $btnDrainAll, $btnLogoffAll, $btnDrainLogoff)) {
        $b.Enabled = $On
    }
}

function Update-Badges {
    $lblHostBadge.Text    = "$($lvHosts.Items.Count)"
    $lblSessionBadge.Text = "$($lvSessions.Items.Count)"
}

# ── Data ───────────────────────────────────────────────────────────────────────

function Get-Broker { "$env:COMPUTERNAME.$((Get-CimInstance Win32_ComputerSystem).Domain)" }

function Load-Hosts {
    $lvHosts.Items.Clear()
    $broker = $txtBroker.Text.Trim()
    try {
        $cols = Get-RDSessionCollection -ConnectionBroker $broker |
                Select-Object -ExpandProperty CollectionName
        foreach ($c in $cols) {
            Get-RDSessionHost -ConnectionBroker $broker -CollectionName $c |
            Select-Object SessionHost, NewConnectionAllowed |
            ForEach-Object {
                $item = [System.Windows.Forms.ListViewItem]::new($_.SessionHost)
                $item.SubItems.Add($_.NewConnectionAllowed) | Out-Null
                $item.Tag       = $_.SessionHost
                $item.ForeColor = switch -Regex ($_.NewConnectionAllowed) {
                    "^Yes$"              { $SUCCESS }
                    "Reboot|No"         { $DANGER  }
                    default             { $WARNING }
                }
                $item.Font = $fUI
                $lvHosts.Items.Add($item) | Out-Null
            }
        }
        Write-Log "Loaded $($lvHosts.Items.Count) session host(s)." $SUCCESS
    } catch {
        Write-Log "Error loading hosts: $_" $DANGER
    }
}

function Load-Sessions {
    $lvSessions.Items.Clear()
    $broker = $txtBroker.Text.Trim()
    try {
        Get-RDUserSession -ConnectionBroker $broker |
        ForEach-Object {
            $item = [System.Windows.Forms.ListViewItem]::new($_.UserName)
            $item.SubItems.Add($_.HostServer)            | Out-Null
            $item.SubItems.Add("$($_.UnifiedSessionId)") | Out-Null
            $item.SubItems.Add($_.SessionState)          | Out-Null
            $item.Tag       = $_
            $item.ForeColor = $TEXT
            $item.Font      = $fUI
            $lvSessions.Items.Add($item) | Out-Null
        }
        Write-Log "Loaded $($lvSessions.Items.Count) active session(s)." $SUCCESS
    } catch {
        Write-Log "Error loading sessions: $_" $DANGER
    }
}

function Invoke-Refresh {
    Set-UI $false
    Write-Log "Refreshing  —  broker: $($txtBroker.Text.Trim())" $ACCENT
    Load-Hosts
    Load-Sessions
    Update-Badges
    Set-UI $true
}

# ── Actions ────────────────────────────────────────────────────────────────────

function Confirm-Action([string]$Title, [string]$Body) {
    $r = [System.Windows.Forms.MessageBox]::Show(
        $Body, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    return $r -eq "Yes"
}

function Invoke-DrainAll {
    if ($lvHosts.Items.Count -eq 0) { Write-Log "No hosts to drain." $WARNING; return }
    if (-not (Confirm-Action "Drain All Hosts" `
        "Set all $($lvHosts.Items.Count) session host(s) to drain mode?`n`nNo new connections will be accepted until each server reboots.")) {
        Write-Log "Drain cancelled." $MUTED; return
    }
    Set-UI $false
    $broker = $txtBroker.Text.Trim()
    foreach ($item in $lvHosts.Items) {
        Write-Log "Draining  $($item.Tag)…" $ACCENT
        try {
            Set-RDSessionHost -SessionHost $item.Tag -NewConnectionAllowed NotUntilReboot -ConnectionBroker $broker
            $item.SubItems[1].Text = "NotUntilReboot"
            $item.ForeColor        = $DANGER
            Write-Log "Drained   $($item.Tag)" $SUCCESS
        } catch { Write-Log "Failed    $($item.Tag): $_" $DANGER }
    }
    Set-UI $true
    Write-Log "All hosts drained." $SUCCESS
}

function Invoke-LogoffAll {
    if ($lvSessions.Items.Count -eq 0) { Write-Log "No sessions to log off." $WARNING; return }
    if (-not (Confirm-Action "Logoff All Sessions" `
        "Force-logoff all $($lvSessions.Items.Count) session(s)?`n`nUsers will lose unsaved work immediately.")) {
        Write-Log "Logoff cancelled." $MUTED; return
    }
    Set-UI $false
    $broker = $txtBroker.Text.Trim()
    foreach ($item in $lvSessions.Items) {
        $s = $item.Tag
        Write-Log "Logging off  $($s.UserName)  (session $($s.UnifiedSessionId)  /  $($s.HostServer))…" $ACCENT
        try {
            Invoke-RDUserLogoff -HostServer $s.HostServer -UnifiedSessionID $s.UnifiedSessionId `
                                -ConnectionBroker $broker -Force
            Write-Log "Logged off   $($s.UserName)" $SUCCESS
        } catch { Write-Log "Failed       $($s.UserName): $_" $DANGER }
    }
    Set-UI $true
    Load-Sessions; Update-Badges
}

function Invoke-DrainAndLogoff {
    $hc = $lvHosts.Items.Count; $sc = $lvSessions.Items.Count
    if ($hc -eq 0 -and $sc -eq 0) { Write-Log "Nothing to do." $WARNING; return }
    if (-not (Confirm-Action "Drain + Logoff All" `
        "Drain $hc host(s) AND force-logoff $sc session(s)?`n`nAll active user sessions will be terminated immediately.")) {
        Write-Log "Operation cancelled." $MUTED; return
    }
    Set-UI $false
    $broker = $txtBroker.Text.Trim()
    Write-Log "━━  Draining hosts  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" $PURPLE
    foreach ($item in $lvHosts.Items) {
        Write-Log "Draining  $($item.Tag)…" $ACCENT
        try {
            Set-RDSessionHost -SessionHost $item.Tag -NewConnectionAllowed NotUntilReboot -ConnectionBroker $broker
            $item.SubItems[1].Text = "NotUntilReboot"; $item.ForeColor = $DANGER
            Write-Log "Drained   $($item.Tag)" $SUCCESS
        } catch { Write-Log "Failed    $($item.Tag): $_" $DANGER }
    }
    Write-Log "━━  Logging off sessions  ━━━━━━━━━━━━━━━━━━━━━━━" $PURPLE
    foreach ($item in $lvSessions.Items) {
        $s = $item.Tag
        Write-Log "Logging off  $($s.UserName)…" $ACCENT
        try {
            Invoke-RDUserLogoff -HostServer $s.HostServer -UnifiedSessionID $s.UnifiedSessionId `
                                -ConnectionBroker $broker -Force
            Write-Log "Logged off   $($s.UserName)" $SUCCESS
        } catch { Write-Log "Failed       $($s.UserName): $_" $DANGER }
    }
    Set-UI $true
    Load-Sessions; Update-Badges
    Write-Log "━━  Complete  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" $SUCCESS
}

# ══════════════════════════════════════════════════════════════════════════════
# FORM
# ══════════════════════════════════════════════════════════════════════════════

$form               = [System.Windows.Forms.Form]::new()
$form.Text          = "RDS Drain and Logoff"
$form.Size          = [System.Drawing.Size]::new(980, 720)
$form.MinimumSize   = [System.Drawing.Size]::new(860, 640)
$form.StartPosition = "CenterScreen"
$form.BackColor     = $BG
$form.Font          = $fUI
$form.Icon          = [System.Drawing.SystemIcons]::Shield

# ── Gradient header ────────────────────────────────────────────────────────────

$pnlHeader          = [System.Windows.Forms.Panel]::new()
$pnlHeader.Dock     = "Top"
$pnlHeader.Height   = 72
$pnlHeader.Add_Paint({
    param($s, $e)
    $r  = [System.Drawing.Rectangle]::new(0, 0, $s.Width, $s.Height)
    $gb = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
        $r, $PURPLE, $ACCENT,
        [System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
    $e.Graphics.FillRectangle($gb, $r)
    $gb.Dispose()
})

$lblTitle           = New-Label "RDS Drain and Logoff" $fTitle $TEXT
$lblTitle.Location  = [System.Drawing.Point]::new(20, 12)

$lblSub             = New-Label "Yuval Grimblat  ·  Network Security Solutions Architect & PM  ·  Mornex LTD 2026" $fSub `
                        ([System.Drawing.Color]::FromArgb(200, 220, 255))
$lblSub.Location    = [System.Drawing.Point]::new(22, 42)

$lnk                = [System.Windows.Forms.LinkLabel]::new()
$lnk.Text           = "LinkedIn"
$lnk.Font           = $fSub
$lnk.LinkColor      = [System.Drawing.Color]::FromArgb(200, 220, 255)
$lnk.ActiveLinkColor= $TEXT
$lnk.BackColor      = [System.Drawing.Color]::Transparent
$lnk.AutoSize       = $true
$lnk.Anchor         = "Top,Right"
$lnk.Add_LinkClicked({ Start-Process "https://www.linkedin.com/in/yuvalgrimblat" })

$pnlHeader.Add_Resize({
    $lnk.Location = [System.Drawing.Point]::new($pnlHeader.Width - $lnk.Width - 20, 26)
})

$pnlHeader.Controls.AddRange(@($lblTitle, $lblSub, $lnk))

# ── Broker bar ─────────────────────────────────────────────────────────────────

$pnlBroker          = [System.Windows.Forms.Panel]::new()
$pnlBroker.Dock     = "Top"
$pnlBroker.Height   = 52
$pnlBroker.BackColor= $SURFACE

$lblBL              = New-Label "Connection Broker" $fBold $MUTED
$lblBL.Location     = [System.Drawing.Point]::new(20, 17)

$txtBroker          = [System.Windows.Forms.TextBox]::new()
$txtBroker.Location = [System.Drawing.Point]::new(176, 14)
$txtBroker.Width    = 380
$txtBroker.Height   = 26
$txtBroker.Font     = $fMono
$txtBroker.BackColor= $CARD
$txtBroker.ForeColor= $TEXT
$txtBroker.BorderStyle = "FixedSingle"
$txtBroker.Text     = (Get-Broker)

$btnRefresh         = New-PillBtn "⟳  Refresh" $ACCENT `
                        ([System.Drawing.Color]::FromArgb(30, 100, 210)) 120
$btnRefresh.Location= [System.Drawing.Point]::new(568, 9)

$pnlBroker.Controls.AddRange(@($lblBL, $txtBroker, $btnRefresh))

# ── Content area ───────────────────────────────────────────────────────────────

$pnlContent         = [System.Windows.Forms.Panel]::new()
$pnlContent.Dock    = "Fill"
$pnlContent.BackColor = $BG
$pnlContent.Padding = [System.Windows.Forms.Padding]::new(16, 12, 16, 0)

# — Hosts card —
$cardHosts          = New-Card
$cardHosts.Size     = [System.Drawing.Size]::new(420, 260)
$cardHosts.Location = [System.Drawing.Point]::new(0, 0)
$cardHosts.Anchor   = "Top,Left,Right"

$lblHostTitle       = New-Label "Session Hosts" $fBold $TEXT
$lblHostTitle.Location = [System.Drawing.Point]::new(16, 14)

$lblHostBadge       = New-Label "0" $fSub $TEXT
$lblHostBadge.BackColor = $ACCENT
$lblHostBadge.AutoSize  = $true
$lblHostBadge.Padding   = [System.Windows.Forms.Padding]::new(6, 2, 6, 2)
$lblHostBadge.Location  = [System.Drawing.Point]::new(128, 12)

$lvHosts            = [DarkListView]::new()
$lvHosts.BackColor  = $CARD
$lvHosts.ForeColor  = $TEXT
$lvHosts.Font       = $fUI
$lvHosts.Location   = [System.Drawing.Point]::new(1, 42)
$lvHosts.Anchor     = "Top,Left,Right,Bottom"
$lvHosts.Columns.Add("Host",   240) | Out-Null
$lvHosts.Columns.Add("Status", 130) | Out-Null

$cardHosts.Controls.AddRange(@($lblHostTitle, $lblHostBadge, $lvHosts))

# — Sessions card —
$cardSessions          = New-Card
$cardSessions.Location = [System.Drawing.Point]::new(440, 0)
$cardSessions.Size     = [System.Drawing.Size]::new(480, 260)
$cardSessions.Anchor   = "Top,Left,Right"

$lblSessTitle          = New-Label "Active Sessions" $fBold $TEXT
$lblSessTitle.Location = [System.Drawing.Point]::new(16, 14)

$lblSessionBadge       = New-Label "0" $fSub $TEXT
$lblSessionBadge.BackColor = $DANGER
$lblSessionBadge.AutoSize  = $true
$lblSessionBadge.Padding   = [System.Windows.Forms.Padding]::new(6, 2, 6, 2)
$lblSessionBadge.Location  = [System.Drawing.Point]::new(142, 12)

$lvSessions            = [DarkListView]::new()
$lvSessions.BackColor  = $CARD
$lvSessions.ForeColor  = $TEXT
$lvSessions.Font       = $fUI
$lvSessions.Location   = [System.Drawing.Point]::new(1, 42)
$lvSessions.Anchor     = "Top,Left,Right,Bottom"
$lvSessions.Columns.Add("User",       140) | Out-Null
$lvSessions.Columns.Add("Host",       150) | Out-Null
$lvSessions.Columns.Add("Session ID",  80) | Out-Null
$lvSessions.Columns.Add("State",       80) | Out-Null

$cardSessions.Controls.AddRange(@($lblSessTitle, $lblSessionBadge, $lvSessions))

# — Action buttons row —
$pnlActions            = [System.Windows.Forms.Panel]::new()
$pnlActions.BackColor  = $BG
$pnlActions.Height     = 52
$pnlActions.Location   = [System.Drawing.Point]::new(0, 276)
$pnlActions.Anchor     = "Top,Left,Right"

$btnDrainAll           = New-PillBtn "⬇  Drain All Hosts" `
    ([System.Drawing.Color]::FromArgb(126, 80, 0)) `
    ([System.Drawing.Color]::FromArgb(160, 105, 10)) 180
$btnDrainAll.Location  = [System.Drawing.Point]::new(0, 8)

$btnLogoffAll          = New-PillBtn "✕  Logoff All Sessions" `
    ([System.Drawing.Color]::FromArgb(110, 30, 28)) `
    ([System.Drawing.Color]::FromArgb(160, 45, 42)) 200
$btnLogoffAll.Location = [System.Drawing.Point]::new(192, 8)

$btnDrainLogoff        = New-PillBtn "⚡  Drain + Logoff All" `
    ([System.Drawing.Color]::FromArgb(80, 30, 140)) `
    ([System.Drawing.Color]::FromArgb(110, 50, 180)) 200
$btnDrainLogoff.Location = [System.Drawing.Point]::new(404, 8)

$pnlActions.Controls.AddRange(@($btnDrainAll, $btnLogoffAll, $btnDrainLogoff))

# — Log card —
$cardLog               = New-Card 10
$cardLog.BackColor     = [System.Drawing.Color]::FromArgb(10, 12, 16)
$cardLog.BorderColor   = $BORDER
$cardLog.Location      = [System.Drawing.Point]::new(0, 336)
$cardLog.Anchor        = "Top,Left,Right,Bottom"

$lblLogTitle           = New-Label "  Status Log" $fBold $MUTED
$lblLogTitle.Location  = [System.Drawing.Point]::new(12, 10)

$logBox                = [System.Windows.Forms.RichTextBox]::new()
$logBox.BackColor      = [System.Drawing.Color]::FromArgb(10, 12, 16)
$logBox.ForeColor      = $TEXT
$logBox.Font           = $fMono
$logBox.BorderStyle    = "None"
$logBox.ReadOnly       = $true
$logBox.ScrollBars     = "Vertical"
$logBox.Location       = [System.Drawing.Point]::new(12, 34)
$logBox.Anchor         = "Top,Left,Right,Bottom"

$cardLog.Controls.AddRange(@($lblLogTitle, $logBox))

$pnlContent.Controls.AddRange(@($cardHosts, $cardSessions, $pnlActions, $cardLog))

# ── Responsive resize ──────────────────────────────────────────────────────────

$form.Add_Resize({
    $w      = $pnlContent.ClientSize.Width - 32   # total width minus 2×16 padding
    $half   = [int](($w - 16) / 2)                # card width with gap

    $cardHosts.Width    = $half
    $cardSessions.Width = $w - $half - 16
    $cardSessions.Left  = $half + 16

    $pnlActions.Width   = $w

    $cardLog.Left   = 0
    $cardLog.Width  = $w
    $cardLog.Top    = $cardHosts.Bottom + $pnlActions.Height + 8
    $cardLog.Height = $pnlContent.ClientSize.Height - $cardLog.Top - 8

    # Keep ListView filling their cards
    $lvHosts.Width    = $cardHosts.ClientSize.Width - 2
    $lvHosts.Height   = $cardHosts.ClientSize.Height - 44
    $lvSessions.Width = $cardSessions.ClientSize.Width - 2
    $lvSessions.Height= $cardSessions.ClientSize.Height - 44

    $logBox.Width  = $cardLog.ClientSize.Width - 24
    $logBox.Height = $cardLog.ClientSize.Height - 44

    $lnk.Location = [System.Drawing.Point]::new($pnlHeader.Width - $lnk.Width - 20, 26)
})

# ── Assemble ───────────────────────────────────────────────────────────────────

$form.Controls.Add($pnlContent)
$form.Controls.Add($pnlBroker)
$form.Controls.Add($pnlHeader)

# ── Events ─────────────────────────────────────────────────────────────────────

$btnRefresh.Add_Click(    { Invoke-Refresh })
$btnDrainAll.Add_Click(   { Invoke-DrainAll })
$btnLogoffAll.Add_Click(  { Invoke-LogoffAll })
$btnDrainLogoff.Add_Click({ Invoke-DrainAndLogoff })

$form.Add_Shown({
    $form.Add_Resize.Invoke($form, [System.EventArgs]::Empty)
    Write-Log "RDS Drain and Logoff Tool  —  Mornex LTD 2026" $PURPLE
    Write-Log "Author: Yuval Grimblat  |  linkedin.com/in/yuvalgrimblat" $MUTED
    Write-Log "Connecting to broker: $($txtBroker.Text.Trim())…" $ACCENT
    Invoke-Refresh
})

[System.Windows.Forms.Application]::Run($form)
