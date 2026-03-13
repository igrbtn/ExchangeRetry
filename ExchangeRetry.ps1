<#
.SYNOPSIS
    ExchangeRetry v0.5 — Exchange Transport Manager GUI.
.DESCRIPTION
    WinForms GUI for monitoring and managing Microsoft Exchange transport.
    All Exchange operations run asynchronously via runspaces (lib/AsyncRunner.ps1).
    GUI never freezes; errors are caught gracefully.
.NOTES
    Version: 0.5.0
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptRoot = $PSScriptRoot
. "$scriptRoot/lib/Core.ps1"
. "$scriptRoot/lib/Diagnostics.ps1"
. "$scriptRoot/lib/Monitoring.ps1"
. "$scriptRoot/lib/AsyncRunner.ps1"

try { Initialize-AppData } catch {}
$script:Settings = try { Get-AppSettings } catch { @{} }

# ─── Script-scope state ─────────────────────────────────────────────────────
$script:Session              = $null
$script:TransportServers     = @()
$script:LastDashboardData    = $null
$script:LastQueueData        = @()
$script:LastMessageData      = @()
$script:LastErrorData        = @()
$script:LastTrackingResults  = @()
$script:LastProtocolData     = @()
$script:LastLogSearchResults = @()
$script:LastHeaderResult     = $null
$script:LastRulesData        = @()
$script:LastCertData         = @()
$script:LastStatisticsData   = @()
$script:LastReportText       = ''

# ═══════════════════════════════════════════════════════════════════════════════
# GUI HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function New-StyledDGV {
    param([switch]$Multi)
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.ReadOnly = $true
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.AutoSizeColumnsMode = 'Fill'
    $dgv.RowHeadersVisible = $false
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245,245,250)
    $dgv.BorderStyle = 'None'
    $dgv.Dock = 'Fill'
    if ($Multi) { $dgv.MultiSelect = $true } else { $dgv.MultiSelect = $false }
    return $dgv
}

function New-Btn {
    param([string]$Text, [int]$W=110, [string]$Color='')
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Width = $W
    $btn.Height = 28
    $btn.FlatStyle = 'Flat'
    $btn.Margin = New-Object System.Windows.Forms.Padding(3,4,3,4)
    if ($Color -eq 'Blue') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(50,100,200)
        $btn.ForeColor = [System.Drawing.Color]::White
    } elseif ($Color -eq 'Red') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(200,50,50)
        $btn.ForeColor = [System.Drawing.Color]::White
    } elseif ($Color -eq 'Green') {
        $btn.BackColor = [System.Drawing.Color]::FromArgb(50,150,80)
        $btn.ForeColor = [System.Drawing.Color]::White
    }
    return $btn
}

function New-ConsoleTextBox {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ScrollBars = 'Both'
    $tb.WordWrap = $false
    $tb.ReadOnly = $true
    $tb.Font = New-Object System.Drawing.Font('Consolas', 11)
    $tb.BackColor = [System.Drawing.Color]::FromArgb(25,25,35)
    $tb.ForeColor = [System.Drawing.Color]::FromArgb(220,220,220)
    $tb.Dock = 'Fill'
    return $tb
}

function New-FlowBar {
    param([int]$H=40)
    $fp = New-Object System.Windows.Forms.FlowLayoutPanel
    $fp.Height = $H
    $fp.Dock = 'Top'
    $fp.FlowDirection = 'LeftToRight'
    $fp.WrapContents = $false
    $fp.Padding = New-Object System.Windows.Forms.Padding(2)
    return $fp
}

function New-BoldLabel {
    param([string]$Text)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(4,8,4,4)
    return $lbl
}

function New-InlineLabel {
    param([string]$Text, [int]$MarginLeft=0)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Text
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding($MarginLeft,8,4,4)
    return $lbl
}

function Show-Export {
    param([object]$Data, [string]$DefaultName='export')
    if (-not $Data -or ($Data | Measure-Object).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No data to export.','Export','OK','Information')
        return
    }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV Files (*.csv)|*.csv|JSON Files (*.json)|*.json'
    $sfd.FileName = "$DefaultName-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            $fmt = if ($sfd.FileName -match '\.json$') { 'JSON' } else { 'CSV' }
            Export-ResultsToFile -Data $Data -FilePath $sfd.FileName -Format $fmt
            Update-StatusBar "Exported to $($sfd.FileName)"
        } catch {
            Update-StatusBar "Export error: $_"
        }
    }
}

function Update-StatusBar {
    param([string]$Text)
    try {
        if ($script:StatusLabel) {
            $script:StatusLabel.Text = $Text
        }
        if ($script:LastActionLabel) {
            $script:LastActionLabel.Text = "Last: $(Get-Date -Format 'HH:mm:ss') $Text"
        }
    } catch {}
}

function Set-DGVData {
    param(
        [System.Windows.Forms.DataGridView]$DGV,
        [array]$Data
    )
    try {
        $DGV.DataSource = $null
        if ($Data -and $Data.Count -gt 0) {
            $dt = New-Object System.Data.DataTable
            $props = $Data[0].PSObject.Properties | ForEach-Object { $_.Name }
            foreach ($p in $props) { [void]$dt.Columns.Add($p) }
            foreach ($item in $Data) {
                $row = $dt.NewRow()
                foreach ($p in $props) { $row[$p] = "$($item.$p)" }
                [void]$dt.Rows.Add($row)
            }
            $DGV.DataSource = $dt
        }
    } catch {
        # Silently handle DGV data binding errors
    }
}

function Get-SelectedScope {
    try {
        if ($script:ScopeCombo -and $script:ScopeCombo.SelectedIndex -gt 0) {
            return $script:ScopeCombo.SelectedItem.ToString()
        }
    } catch {}
    return $null
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN GUI FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

function Show-ExchangeRetryGUI {

    # ─── Form ────────────────────────────────────────────────────────────────
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'ExchangeRetry v0.5 — Exchange Transport Manager'
    $form.Size = New-Object System.Drawing.Size(1400, 900)
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 700)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $form.StartPosition = 'CenterScreen'
    $form.KeyPreview = $true

    # Restore window size from settings
    try {
        if ($script:Settings.WindowWidth -and $script:Settings.WindowHeight) {
            $form.Size = New-Object System.Drawing.Size([int]$script:Settings.WindowWidth, [int]$script:Settings.WindowHeight)
        }
    } catch {}

    # ─── Top Panel ───────────────────────────────────────────────────────────
    $topPanel = New-FlowBar -H 44
    $topPanel.BackColor = [System.Drawing.Color]::FromArgb(240,240,245)

    $lblServer = New-BoldLabel -Text 'Exchange Server:'
    $txtServer = New-Object System.Windows.Forms.TextBox
    $txtServer.Width = 220
    $txtServer.Height = 24
    $txtServer.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    # Autocomplete from recent servers
    $txtServer.AutoCompleteMode = 'SuggestAppend'
    $txtServer.AutoCompleteSource = 'CustomSource'
    $autoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
    try {
        if ($script:Settings.RecentServers) {
            foreach ($s in $script:Settings.RecentServers) { [void]$autoComplete.Add($s) }
        }
    } catch {}
    $txtServer.AutoCompleteCustomSource = $autoComplete
    try { if ($script:Settings.LastServer) { $txtServer.Text = $script:Settings.LastServer } } catch {}

    $btnConnect = New-Btn -Text 'Connect' -W 90 -Color 'Blue'
    $btnDisconnect = New-Btn -Text 'Disconnect' -W 90 -Color 'Red'
    $btnDisconnect.Visible = $false
    $lblConnStatus = New-InlineLabel -Text 'Not connected' -MarginLeft 6
    $lblConnStatus.ForeColor = [System.Drawing.Color]::Gray

    $lblScope = New-BoldLabel -Text 'Scope:'
    $script:ScopeCombo = New-Object System.Windows.Forms.ComboBox
    $script:ScopeCombo.Width = 200
    $script:ScopeCombo.DropDownStyle = 'DropDownList'
    $script:ScopeCombo.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    [void]$script:ScopeCombo.Items.Add('(All Servers)')
    $script:ScopeCombo.SelectedIndex = 0

    $chkAutoRefresh = New-Object System.Windows.Forms.CheckBox
    $chkAutoRefresh.Text = 'Auto-refresh'
    $chkAutoRefresh.AutoSize = $true
    $chkAutoRefresh.Margin = New-Object System.Windows.Forms.Padding(20,8,3,4)

    $topPanel.Controls.AddRange(@($lblServer, $txtServer, $btnConnect, $btnDisconnect, $lblConnStatus, $lblScope, $script:ScopeCombo, $chkAutoRefresh))
    $form.Controls.Add($topPanel)

    # ─── Status Bar ──────────────────────────────────────────────────────────
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $script:StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:StatusLabel.Spring = $true
    $script:StatusLabel.TextAlign = 'MiddleLeft'
    $script:StatusLabel.Text = 'Ready'
    $script:LastActionLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $script:LastActionLabel.Text = ''
    $script:LastActionLabel.Alignment = 'Right'
    [void]$statusStrip.Items.Add($script:StatusLabel)
    [void]$statusStrip.Items.Add($script:LastActionLabel)
    $form.Controls.Add($statusStrip)

    # ─── Job Console Panel (bottom) ──────────────────────────────────────────
    $jobPanel = New-JobConsolePanel -Height 130
    $form.Controls.Add($jobPanel)

    # ─── Tab Control ─────────────────────────────────────────────────────────
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 1: DASHBOARD
    # ═══════════════════════════════════════════════════════════════════════════
    $tabDash = New-Object System.Windows.Forms.TabPage
    $tabDash.Text = 'Dashboard'

    $dashToolbar = New-FlowBar -H 38
    $btnDashRefresh = New-Btn -Text 'Refresh' -W 90
    $dashToolbar.Controls.Add($btnDashRefresh)
    $tabDash.Controls.Add($dashToolbar)

    $dashSplit = New-Object System.Windows.Forms.SplitContainer
    $dashSplit.Dock = 'Fill'
    $dashSplit.Orientation = 'Horizontal'
    $dashSplit.SplitterDistance = 450

    $txtDashboard = New-ConsoleTextBox
    $dashSplit.Panel1.Controls.Add($txtDashboard)

    $alertPanel = New-Object System.Windows.Forms.Panel
    $alertPanel.Dock = 'Fill'
    $lblAlerts = New-BoldLabel -Text 'ALERTS: All clear'
    $lblAlerts.Dock = 'Top'
    $lblAlerts.ForeColor = [System.Drawing.Color]::Green
    $lblAlerts.Padding = New-Object System.Windows.Forms.Padding(4)
    $txtAlerts = New-ConsoleTextBox
    $alertPanel.Controls.Add($txtAlerts)
    $alertPanel.Controls.Add($lblAlerts)
    $dashSplit.Panel2.Controls.Add($alertPanel)

    $tabDash.Controls.Add($dashSplit)

    $refreshDashboard = {
        Update-StatusBar 'Refreshing dashboard...'
        $scopeVal = Get-SelectedScope
        Start-AsyncJob -Name 'Dashboard' -Form $form -ScriptBlock {
            param($Server)
            $data = Get-DashboardData -Server $Server
            return $data
        } -Parameters @{ Server = $scopeVal } -OnComplete {
            param($result)
            try {
                $script:LastDashboardData = $result
                if ($result) {
                    $txtDashboard.Text = ($result | Out-String)
                }
                $qd = if ($script:LastQueueData) { $script:LastQueueData } else { @() }
                $alerts = Test-TransportAlerts -QueueData $qd
                $alertText = Format-AlertText -Alerts $alerts
                $txtAlerts.Text = $alertText

                $maxLevel = 'Info'
                foreach ($a in $alerts) {
                    if ($a.Level -eq 'Critical') { $maxLevel = 'Critical'; break }
                    if ($a.Level -eq 'Warning') { $maxLevel = 'Warning' }
                }
                switch ($maxLevel) {
                    'Critical' { $lblAlerts.Text = 'ALERTS: CRITICAL'; $lblAlerts.ForeColor = [System.Drawing.Color]::Red }
                    'Warning'  { $lblAlerts.Text = 'ALERTS: Warning';  $lblAlerts.ForeColor = [System.Drawing.Color]::FromArgb(200,150,0) }
                    default    { $lblAlerts.Text = 'ALERTS: All clear'; $lblAlerts.ForeColor = [System.Drawing.Color]::Green }
                }
                Update-StatusBar 'Dashboard refreshed'
            } catch {
                Update-StatusBar "Dashboard UI error: $_"
            }
        } -OnError {
            param($err)
            $txtDashboard.Text = "Error: $err"
            Update-StatusBar "Dashboard error: $err"
        }
    }
    $btnDashRefresh.Add_Click($refreshDashboard)

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 2: QUEUES (3-panel vertical split)
    # ═══════════════════════════════════════════════════════════════════════════
    $tabQueues = New-Object System.Windows.Forms.TabPage
    $tabQueues.Text = 'Queues'

    $queueSplitOuter = New-Object System.Windows.Forms.SplitContainer
    $queueSplitOuter.Dock = 'Fill'
    $queueSplitOuter.Orientation = 'Horizontal'
    $queueSplitOuter.SplitterDistance = 280

    $queueSplitInner = New-Object System.Windows.Forms.SplitContainer
    $queueSplitInner.Dock = 'Fill'
    $queueSplitInner.Orientation = 'Horizontal'
    $queueSplitInner.SplitterDistance = 250

    # --- Top: Queue grid ---
    $queueTopPanel = New-Object System.Windows.Forms.Panel
    $queueTopPanel.Dock = 'Fill'

    $queueToolbar = New-FlowBar -H 38
    $btnQueueRefresh = New-Btn -Text 'Refresh' -W 80
    $btnRetryQueue = New-Btn -Text 'Retry Queue' -W 100
    $txtQueueFilter = New-Object System.Windows.Forms.TextBox
    $txtQueueFilter.Width = 180
    $txtQueueFilter.Height = 24
    $txtQueueFilter.Margin = New-Object System.Windows.Forms.Padding(10,6,3,4)
    $lblQueueFilter = New-InlineLabel -Text 'Filter:'
    $queueToolbar.Controls.AddRange(@($btnQueueRefresh, $btnRetryQueue, $lblQueueFilter, $txtQueueFilter))

    $dgvQueues = New-StyledDGV

    # Queue context menu
    $ctxQueue = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxQueueRetry = $ctxQueue.Items.Add('Retry Queue')
    $ctxQueueSuspend = $ctxQueue.Items.Add('Suspend Queue')
    $ctxQueueCopyId = $ctxQueue.Items.Add('Copy Identity')
    $dgvQueues.ContextMenuStrip = $ctxQueue

    # Queue row color coding
    $dgvQueues.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $statusCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Status') { $statusCol = $c.Index; break } }
            if ($null -ne $statusCol) {
                $val = "$($row.Cells[$statusCol].Value)"
                if ($val -match 'Retry')     { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200) }
                elseif ($val -match 'Suspended') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,180) }
                elseif ($val -match 'Active|Ready') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
            }
        } catch {}
    })

    $queueTopPanel.Controls.Add($dgvQueues)
    $queueTopPanel.Controls.Add($queueToolbar)
    $queueSplitOuter.Panel1.Controls.Add($queueTopPanel)

    # --- Middle: Messages grid ---
    $msgPanel = New-Object System.Windows.Forms.Panel
    $msgPanel.Dock = 'Fill'

    $msgToolbar = New-FlowBar -H 38
    $btnMsgRetry = New-Btn -Text 'Retry Selected' -W 110
    $btnMsgSuspend = New-Btn -Text 'Suspend Selected' -W 120
    $btnMsgRemove = New-Btn -Text 'Remove Selected' -W 115 -Color 'Red'
    $chkNDR = New-Object System.Windows.Forms.CheckBox
    $chkNDR.Text = 'NDR'
    $chkNDR.AutoSize = $true
    $chkNDR.Margin = New-Object System.Windows.Forms.Padding(10,8,3,4)
    $msgToolbar.Controls.AddRange(@($btnMsgRetry, $btnMsgSuspend, $btnMsgRemove, $chkNDR))

    $dgvMessages = New-StyledDGV -Multi

    # Messages context menu
    $ctxMsg = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxMsgRetry = $ctxMsg.Items.Add('Retry')
    $ctxMsgSuspend = $ctxMsg.Items.Add('Suspend')
    $ctxMsgRemove = $ctxMsg.Items.Add('Remove')
    $ctxMsgCopyId = $ctxMsg.Items.Add('Copy MessageId')
    $ctxMsgTrack = $ctxMsg.Items.Add('Track in Message Tracking')
    $dgvMessages.ContextMenuStrip = $ctxMsg

    # Message row color coding
    $dgvMessages.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $statusCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Status') { $statusCol = $c.Index; break } }
            if ($null -ne $statusCol) {
                $val = "$($row.Cells[$statusCol].Value)"
                if ($val -match 'Retry|Suspended') {
                    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,200,200)
                }
            }
        } catch {}
    })

    $msgPanel.Controls.Add($dgvMessages)
    $msgPanel.Controls.Add($msgToolbar)
    $queueSplitInner.Panel1.Controls.Add($msgPanel)

    # --- Bottom: Recent Errors ---
    $errorPanel = New-Object System.Windows.Forms.Panel
    $errorPanel.Dock = 'Fill'

    $errorToolbar = New-FlowBar -H 38
    $btnErrorRefresh = New-Btn -Text 'Refresh Errors' -W 110
    $btnErrorExport = New-Btn -Text 'Export...' -W 80
    $lblErrorSummary = New-InlineLabel -Text 'FAIL: 0 | DEFER: 0 | DSN: 0' -MarginLeft 10
    $errorToolbar.Controls.AddRange(@($btnErrorRefresh, $btnErrorExport, $lblErrorSummary))

    $dgvErrors = New-StyledDGV

    # Error row color coding
    $dgvErrors.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $eventCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'EventId') { $eventCol = $c.Index; break } }
            if ($null -ne $eventCol) {
                $val = "$($row.Cells[$eventCol].Value)"
                if ($val -match 'FAIL')  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,180,180) }
                elseif ($val -match 'DEFER') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,180) }
                elseif ($val -match 'DSN')   { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,210,160) }
            }
        } catch {}
    })

    $errorPanel.Controls.Add($dgvErrors)
    $errorPanel.Controls.Add($errorToolbar)
    $queueSplitInner.Panel2.Controls.Add($errorPanel)

    $queueSplitOuter.Panel2.Controls.Add($queueSplitInner)
    $tabQueues.Controls.Add($queueSplitOuter)

    # --- Queue async actions ---
    $refreshQueues = {
        Update-StatusBar 'Refreshing queues...'
        $scopeVal = Get-SelectedScope
        $filterText = $txtQueueFilter.Text
        Start-AsyncJob -Name 'Queues' -Form $form -ScriptBlock {
            param($Server)
            return @(Get-ExchangeQueues -Server $Server)
        } -Parameters @{ Server = $scopeVal } -OnComplete {
            param($result)
            try {
                $script:LastQueueData = @($result)
                $filtered = $script:LastQueueData
                if ($filterText) {
                    $f = $filterText
                    $filtered = $filtered | Where-Object { ($_ | Out-String) -match [regex]::Escape($f) }
                }
                Set-DGVData -DGV $dgvQueues -Data $filtered
                Update-StatusBar "Queues: $($filtered.Count) loaded"
            } catch {
                Update-StatusBar "Queue UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Queue error: $err"
        }
    }
    $btnQueueRefresh.Add_Click($refreshQueues)
    $txtQueueFilter.Add_TextChanged({
        # Filter is local-only (no async needed), just re-filter cached data
        try {
            $filtered = $script:LastQueueData
            if ($txtQueueFilter.Text) {
                $f = $txtQueueFilter.Text
                $filtered = $filtered | Where-Object { ($_ | Out-String) -match [regex]::Escape($f) }
            }
            Set-DGVData -DGV $dgvQueues -Data $filtered
        } catch {}
    })

    $btnRetryQueue.Add_Click({
        try {
            if ($dgvQueues.SelectedRows.Count -eq 0) { return }
            $identity = "$($dgvQueues.SelectedRows[0].Cells['Identity'].Value)"
            $confirm = [System.Windows.Forms.MessageBox]::Show("Retry queue '$identity'?", 'Confirm', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "RetryQueue $identity" -Form $form -ScriptBlock {
                    param($Identity)
                    Invoke-QueueRetry -Identity $Identity
                } -Parameters @{ Identity = $identity } -OnComplete {
                    param($r)
                    try { Write-OperatorLog -Action 'RetryQueue' -Target $identity } catch {}
                    Update-StatusBar "Queue '$identity' retried"
                    & $refreshQueues
                } -OnError {
                    param($err)
                    Update-StatusBar "Retry error: $err"
                }
            }
        } catch { Update-StatusBar "Retry queue error: $_" }
    })
    $ctxQueueRetry.Add_Click({
        $btnRetryQueue.PerformClick()
    })

    $ctxQueueSuspend.Add_Click({
        try {
            if ($dgvQueues.SelectedRows.Count -eq 0) { return }
            $identity = "$($dgvQueues.SelectedRows[0].Cells['Identity'].Value)"
            $confirm = [System.Windows.Forms.MessageBox]::Show("Suspend queue '$identity'?", 'Confirm', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "SuspendQueue $identity" -Form $form -ScriptBlock {
                    param($Identity)
                    Suspend-Queue -Identity $Identity -Confirm:$false
                } -Parameters @{ Identity = $identity } -OnComplete {
                    param($r)
                    try { Write-OperatorLog -Action 'SuspendQueue' -Target $identity } catch {}
                    Update-StatusBar "Queue '$identity' suspended"
                    & $refreshQueues
                } -OnError {
                    param($err)
                    Update-StatusBar "Suspend error: $err"
                }
            }
        } catch { Update-StatusBar "Suspend queue error: $_" }
    })

    $ctxQueueCopyId.Add_Click({
        try {
            if ($dgvQueues.SelectedRows.Count -gt 0) {
                $identity = "$($dgvQueues.SelectedRows[0].Cells['Identity'].Value)"
                [System.Windows.Forms.Clipboard]::SetText($identity)
                Update-StatusBar "Copied: $identity"
            }
        } catch {}
    })

    # Double-click queue -> load messages (async)
    $dgvQueues.Add_CellDoubleClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $identity = "$($dgvQueues.Rows[$e.RowIndex].Cells['Identity'].Value)"
            Update-StatusBar "Loading messages for $identity..."
            Start-AsyncJob -Name "Messages $identity" -Form $form -ScriptBlock {
                param($Identity)
                return @(Get-QueueMessages -Identity $Identity)
            } -Parameters @{ Identity = $identity } -OnComplete {
                param($result)
                try {
                    $script:LastMessageData = @($result)
                    Set-DGVData -DGV $dgvMessages -Data $script:LastMessageData
                    Update-StatusBar "Messages: $($script:LastMessageData.Count) loaded"
                } catch {
                    Update-StatusBar "Message UI error: $_"
                }
            } -OnError {
                param($err)
                Update-StatusBar "Message load error: $err"
            }
        } catch {}
    })

    # --- Message actions (async) ---
    $getSelectedMessageIds = {
        $ids = @()
        try {
            foreach ($row in $dgvMessages.SelectedRows) {
                $mid = "$($row.Cells['MessageId'].Value)"
                if ($mid) { $ids += $mid }
            }
        } catch {}
        return $ids
    }

    $btnMsgRetry.Add_Click({
        try {
            $ids = & $getSelectedMessageIds
            if ($ids.Count -eq 0) { return }
            $confirm = [System.Windows.Forms.MessageBox]::Show("Retry $($ids.Count) message(s)?", 'Confirm', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "RetryMessages ($($ids.Count))" -Form $form -ScriptBlock {
                    param($MessageIds)
                    foreach ($id in $MessageIds) {
                        try { Invoke-MessageRetry -MessageId $id } catch {}
                    }
                } -Parameters @{ MessageIds = $ids } -OnComplete {
                    param($r)
                    foreach ($id in $ids) {
                        try { Write-OperatorLog -Action 'RetryMessage' -Target $id } catch {}
                    }
                    Update-StatusBar "Retried $($ids.Count) messages"
                } -OnError {
                    param($err)
                    Update-StatusBar "Retry error: $err"
                }
            }
        } catch { Update-StatusBar "Retry messages error: $_" }
    })
    $ctxMsgRetry.Add_Click({ $btnMsgRetry.PerformClick() })

    $btnMsgSuspend.Add_Click({
        try {
            $ids = & $getSelectedMessageIds
            if ($ids.Count -eq 0) { return }
            $confirm = [System.Windows.Forms.MessageBox]::Show("Suspend $($ids.Count) message(s)?", 'Confirm', 'YesNo', 'Question')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "SuspendMessages ($($ids.Count))" -Form $form -ScriptBlock {
                    param($MessageIds)
                    foreach ($id in $MessageIds) {
                        try { Suspend-QueueMessages -MessageId $id } catch {}
                    }
                } -Parameters @{ MessageIds = $ids } -OnComplete {
                    param($r)
                    foreach ($id in $ids) {
                        try { Write-OperatorLog -Action 'SuspendMessage' -Target $id } catch {}
                    }
                    Update-StatusBar "Suspended $($ids.Count) messages"
                } -OnError {
                    param($err)
                    Update-StatusBar "Suspend error: $err"
                }
            }
        } catch { Update-StatusBar "Suspend messages error: $_" }
    })
    $ctxMsgSuspend.Add_Click({ $btnMsgSuspend.PerformClick() })

    $btnMsgRemove.Add_Click({
        try {
            $ids = & $getSelectedMessageIds
            if ($ids.Count -eq 0) { return }
            $ndr = $chkNDR.Checked
            $msg = "Remove $($ids.Count) message(s)?"
            if ($ndr) { $msg += " (with NDR)" }
            $confirm = [System.Windows.Forms.MessageBox]::Show($msg, 'Confirm', 'YesNo', 'Warning')
            if ($confirm -eq 'Yes') {
                Start-AsyncJob -Name "RemoveMessages ($($ids.Count))" -Form $form -ScriptBlock {
                    param($MessageIds, $WithNDR)
                    foreach ($id in $MessageIds) {
                        try { Remove-QueueMessages -MessageId $id -WithNDR:$WithNDR } catch {}
                    }
                } -Parameters @{ MessageIds = $ids; WithNDR = $ndr } -OnComplete {
                    param($r)
                    foreach ($id in $ids) {
                        try { Write-OperatorLog -Action 'RemoveMessage' -Target $id -Details "NDR=$ndr" } catch {}
                    }
                    Update-StatusBar "Removed $($ids.Count) messages"
                } -OnError {
                    param($err)
                    Update-StatusBar "Remove error: $err"
                }
            }
        } catch { Update-StatusBar "Remove messages error: $_" }
    })
    $ctxMsgRemove.Add_Click({ $btnMsgRemove.PerformClick() })

    $ctxMsgCopyId.Add_Click({
        try {
            if ($dgvMessages.SelectedRows.Count -gt 0) {
                $mid = "$($dgvMessages.SelectedRows[0].Cells['MessageId'].Value)"
                [System.Windows.Forms.Clipboard]::SetText($mid)
                Update-StatusBar "Copied: $mid"
            }
        } catch {}
    })

    # Track in Message Tracking context menu
    $ctxMsgTrack.Add_Click({
        try {
            if ($dgvMessages.SelectedRows.Count -gt 0) {
                $mid = "$($dgvMessages.SelectedRows[0].Cells['MessageId'].Value)"
                $txtTrackMsgId.Text = $mid
                $tabs.SelectedTab = $tabTracking
            }
        } catch {}
    })

    # --- Error actions (async) ---
    $refreshErrors = {
        Update-StatusBar 'Refreshing errors...'
        $scopeVal = Get-SelectedScope
        Start-AsyncJob -Name 'RecentErrors' -Form $form -ScriptBlock {
            param($Server)
            return @(Get-RecentErrors -Server $Server)
        } -Parameters @{ Server = $scopeVal } -OnComplete {
            param($result)
            try {
                $script:LastErrorData = @($result)
                Set-DGVData -DGV $dgvErrors -Data $script:LastErrorData

                $failCount = ($result | Where-Object { $_.EventId -match 'FAIL' } | Measure-Object).Count
                $deferCount = ($result | Where-Object { $_.EventId -match 'DEFER' } | Measure-Object).Count
                $dsnCount = ($result | Where-Object { $_.EventId -match 'DSN' } | Measure-Object).Count
                $lblErrorSummary.Text = "FAIL: $failCount | DEFER: $deferCount | DSN: $dsnCount"
                Update-StatusBar "Errors: $($result.Count) loaded"
            } catch {
                Update-StatusBar "Error UI: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Error refresh: $err"
        }
    }
    $btnErrorRefresh.Add_Click($refreshErrors)
    $btnErrorExport.Add_Click({ Show-Export -Data $script:LastErrorData -DefaultName 'errors' })

    # Double-click error -> Message Tracking
    $dgvErrors.Add_CellDoubleClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $midCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'MessageId') { $midCol = $c.Index; break } }
            if ($null -ne $midCol) {
                $mid = "$($dgvErrors.Rows[$e.RowIndex].Cells[$midCol].Value)"
                if ($mid) {
                    $txtTrackMsgId.Text = $mid
                    $tabs.SelectedTab = $tabTracking
                }
            }
        } catch {}
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 3: MESSAGE TRACKING
    # ═══════════════════════════════════════════════════════════════════════════
    $tabTracking = New-Object System.Windows.Forms.TabPage
    $tabTracking.Text = 'Message Tracking'

    $trackPanel = New-Object System.Windows.Forms.Panel
    $trackPanel.Dock = 'Fill'

    # Search bar row 1
    $trackBar1 = New-FlowBar -H 38
    $lblTrkMsgId = New-InlineLabel -Text 'Message-ID:'
    $txtTrackMsgId = New-Object System.Windows.Forms.TextBox
    $txtTrackMsgId.Width = 260
    $txtTrackMsgId.Height = 24
    $txtTrackMsgId.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    $lblTrkSender = New-InlineLabel -Text 'Sender:'
    $txtTrackSender = New-Object System.Windows.Forms.TextBox
    $txtTrackSender.Width = 180
    $txtTrackSender.Height = 24
    $txtTrackSender.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    $lblTrkRecip = New-InlineLabel -Text 'Recipient:'
    $txtTrackRecip = New-Object System.Windows.Forms.TextBox
    $txtTrackRecip.Width = 180
    $txtTrackRecip.Height = 24
    $txtTrackRecip.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $trackBar1.Controls.AddRange(@($lblTrkMsgId, $txtTrackMsgId, $lblTrkSender, $txtTrackSender, $lblTrkRecip, $txtTrackRecip))

    # Search bar row 2
    $trackBar2 = New-FlowBar -H 38
    $lblTrkStart = New-InlineLabel -Text 'Start:'
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = 'Custom'
    $dtpStart.CustomFormat = 'yyyy-MM-dd HH:mm'
    $dtpStart.Width = 150
    $dtpStart.Value = (Get-Date).AddDays(-1)
    $dtpStart.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    $lblTrkEnd = New-InlineLabel -Text 'End:'
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = 'Custom'
    $dtpEnd.CustomFormat = 'yyyy-MM-dd HH:mm'
    $dtpEnd.Width = 150
    $dtpEnd.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    $lblTrkEvent = New-InlineLabel -Text 'EventId:'
    $cmbEventId = New-Object System.Windows.Forms.ComboBox
    $cmbEventId.DropDownStyle = 'DropDownList'
    $cmbEventId.Width = 120
    $cmbEventId.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    @('All','RECEIVE','SEND','DELIVER','SUBMIT','FAIL','DSN','DEFER','EXPAND','REDIRECT','RESOLVE','TRANSFER','POISONMESSAGE') | ForEach-Object { [void]$cmbEventId.Items.Add($_) }
    $cmbEventId.SelectedIndex = 0
    $lblTrkSubject = New-InlineLabel -Text 'Subject:'
    $txtTrackSubject = New-Object System.Windows.Forms.TextBox
    $txtTrackSubject.Width = 150
    $txtTrackSubject.Height = 24
    $txtTrackSubject.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    $btnTrackSearch = New-Btn -Text 'Search' -W 80 -Color 'Blue'
    $trackBar2.Controls.AddRange(@($lblTrkStart, $dtpStart, $lblTrkEnd, $dtpEnd, $lblTrkEvent, $cmbEventId, $lblTrkSubject, $txtTrackSubject, $btnTrackSearch))

    # Results grid
    $dgvTracking = New-StyledDGV -Multi

    # Tracking row color coding by EventId
    $dgvTracking.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $eventCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'EventId') { $eventCol = $c.Index; break } }
            if ($null -ne $eventCol) {
                $val = "$($row.Cells[$eventCol].Value)"
                switch -Regex ($val) {
                    'DELIVER' { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
                    'FAIL'    { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,180,180) }
                    'DEFER'   { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,180) }
                    'SEND'    { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,240,255) }
                    'RECEIVE' { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230,230,255) }
                    'DSN'     { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,210,160) }
                }
            }
        } catch {}
    })

    # Tracking context menu
    $ctxTracking = New-Object System.Windows.Forms.ContextMenuStrip
    $ctxTrkPath = $ctxTracking.Items.Add('Show Message Path')
    $ctxTrkCopyMsgId = $ctxTracking.Items.Add('Copy MessageId')
    $ctxTrkCopySender = $ctxTracking.Items.Add('Copy Sender')
    $ctxTrkExport = $ctxTracking.Items.Add('Export Selected')
    $dgvTracking.ContextMenuStrip = $ctxTracking

    # Bottom bar
    $trackBottom = New-FlowBar -H 38
    $btnTrackExport = New-Btn -Text 'Export...' -W 80
    $btnShowPath = New-Btn -Text 'Show Message Path' -W 140
    $btnCrossTrace = New-Btn -Text 'Cross-Server Trace' -W 140
    $lblTrackCount = New-InlineLabel -Text '0 results' -MarginLeft 10
    $trackBottom.Dock = 'Bottom'
    $trackBottom.Controls.AddRange(@($btnTrackExport, $btnShowPath, $btnCrossTrace, $lblTrackCount))

    $trackPanel.Controls.Add($dgvTracking)
    $trackPanel.Controls.Add($trackBottom)
    $trackPanel.Controls.Add($trackBar2)
    $trackPanel.Controls.Add($trackBar1)
    $tabTracking.Controls.Add($trackPanel)

    # --- Tracking actions (async) ---
    $doTrackSearch = {
        Update-StatusBar 'Searching message tracking logs...'
        $searchParams = @{
            Start = $dtpStart.Value
            End   = $dtpEnd.Value
        }
        if ($txtTrackMsgId.Text)    { $searchParams['MessageId'] = $txtTrackMsgId.Text }
        if ($txtTrackSender.Text)   { $searchParams['Sender'] = $txtTrackSender.Text }
        if ($txtTrackRecip.Text)    { $searchParams['Recipients'] = $txtTrackRecip.Text }
        if ($txtTrackSubject.Text)  { $searchParams['MessageSubject'] = $txtTrackSubject.Text }
        if ($cmbEventId.SelectedItem -and $cmbEventId.SelectedItem -ne 'All') {
            $searchParams['EventId'] = $cmbEventId.SelectedItem.ToString()
        }
        $scopeVal = Get-SelectedScope
        if ($scopeVal) { $searchParams['Server'] = $scopeVal }

        Start-AsyncJob -Name 'MessageTracking' -Form $form -ScriptBlock {
            param($Params)
            return @(Trace-ExchangeMessage @Params)
        } -Parameters @{ Params = $searchParams } -OnComplete {
            param($result)
            try {
                $script:LastTrackingResults = @($result)
                Set-DGVData -DGV $dgvTracking -Data $script:LastTrackingResults
                $lblTrackCount.Text = "$($script:LastTrackingResults.Count) results"
                Update-StatusBar "Tracking: $($script:LastTrackingResults.Count) results"
            } catch {
                Update-StatusBar "Tracking UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Tracking error: $err"
        }
    }
    $btnTrackSearch.Add_Click($doTrackSearch)
    $btnTrackExport.Add_Click({ Show-Export -Data $script:LastTrackingResults -DefaultName 'tracking' })

    $showMessagePath = {
        try {
            if ($dgvTracking.SelectedRows.Count -eq 0) { return }
            $mid = "$($dgvTracking.SelectedRows[0].Cells['MessageId'].Value)"
            if (-not $mid) { return }
            $pathData = $script:LastTrackingResults | Where-Object { $_.MessageId -eq $mid } | Sort-Object Timestamp
            $pathForm = New-Object System.Windows.Forms.Form
            $pathForm.Text = "Message Path: $mid"
            $pathForm.Size = New-Object System.Drawing.Size(800, 500)
            $pathForm.StartPosition = 'CenterParent'
            $pathTxt = New-ConsoleTextBox
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.AppendLine("Message Path for: $mid")
            [void]$sb.AppendLine("=" * 70)
            $step = 1
            foreach ($entry in $pathData) {
                [void]$sb.AppendLine("Step $step`: [$($entry.EventId)] $($entry.Timestamp)")
                [void]$sb.AppendLine("  Server: $($entry.ServerHostname)  Source: $($entry.Source)")
                [void]$sb.AppendLine("  Sender: $($entry.Sender)  Recipients: $($entry.Recipients)")
                if ($entry.PSObject.Properties['ConnectorId'] -and $entry.ConnectorId) {
                    [void]$sb.AppendLine("  Connector: $($entry.ConnectorId)")
                }
                [void]$sb.AppendLine("")
                $step++
            }
            $pathTxt.Text = $sb.ToString()
            $pathForm.Controls.Add($pathTxt)
            $pathForm.ShowDialog()
        } catch { Update-StatusBar "Path error: $_" }
    }
    $btnShowPath.Add_Click($showMessagePath)
    $ctxTrkPath.Add_Click($showMessagePath)

    $btnCrossTrace.Add_Click({
        try {
            if ($dgvTracking.SelectedRows.Count -eq 0) { return }
            $mid = "$($dgvTracking.SelectedRows[0].Cells['MessageId'].Value)"
            if (-not $mid) { return }
            Update-StatusBar "Cross-server trace for $mid..."
            $serverList = @($script:TransportServers | ForEach-Object { $_.Name })
            if ($serverList.Count -eq 0 -and $txtServer.Text) { $serverList = @($txtServer.Text) }
            Start-AsyncJob -Name "CrossTrace $mid" -Form $form -ScriptBlock {
                param($MessageId, $Servers)
                return @(Trace-CrossServerMessage -MessageId $MessageId -Servers $Servers)
            } -Parameters @{ MessageId = $mid; Servers = $serverList } -OnComplete {
                param($result)
                try {
                    $crossForm = New-Object System.Windows.Forms.Form
                    $crossForm.Text = "Cross-Server Trace: $mid"
                    $crossForm.Size = New-Object System.Drawing.Size(900, 600)
                    $crossForm.StartPosition = 'CenterParent'
                    $crossDGV = New-StyledDGV
                    Set-DGVData -DGV $crossDGV -Data @($result)
                    $crossForm.Controls.Add($crossDGV)
                    $crossForm.ShowDialog()
                    Update-StatusBar 'Cross-server trace complete'
                } catch {
                    Update-StatusBar "Cross-trace UI error: $_"
                }
            } -OnError {
                param($err)
                Update-StatusBar "Cross-trace error: $err"
            }
        } catch { Update-StatusBar "Cross-trace error: $_" }
    })

    $ctxTrkCopyMsgId.Add_Click({
        try {
            if ($dgvTracking.SelectedRows.Count -gt 0) {
                $v = "$($dgvTracking.SelectedRows[0].Cells['MessageId'].Value)"
                [System.Windows.Forms.Clipboard]::SetText($v)
            }
        } catch {}
    })
    $ctxTrkCopySender.Add_Click({
        try {
            if ($dgvTracking.SelectedRows.Count -gt 0) {
                $v = "$($dgvTracking.SelectedRows[0].Cells['Sender'].Value)"
                [System.Windows.Forms.Clipboard]::SetText($v)
            }
        } catch {}
    })
    $ctxTrkExport.Add_Click({
        try {
            $selected = @()
            foreach ($row in $dgvTracking.SelectedRows) {
                $obj = @{}
                foreach ($cell in $row.Cells) { $obj[$cell.OwningColumn.Name] = $cell.Value }
                $selected += [PSCustomObject]$obj
            }
            Show-Export -Data $selected -DefaultName 'tracking-selected'
        } catch {}
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 4: PROTOCOL LOGS
    # ═══════════════════════════════════════════════════════════════════════════
    $tabProtocol = New-Object System.Windows.Forms.TabPage
    $tabProtocol.Text = 'Protocol Logs'

    $protoPanel = New-Object System.Windows.Forms.Panel
    $protoPanel.Dock = 'Fill'

    $protoToolbar = New-FlowBar -H 38
    $lblProtoPath = New-InlineLabel -Text 'Log Path:'
    $txtProtoPath = New-Object System.Windows.Forms.TextBox
    $txtProtoPath.Width = 300
    $txtProtoPath.Height = 24
    $txtProtoPath.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnProtoBrowse = New-Btn -Text 'Browse...' -W 80
    $lblProtoFilter = New-InlineLabel -Text 'Filter:' -MarginLeft 10
    $txtProtoFilter = New-Object System.Windows.Forms.TextBox
    $txtProtoFilter.Width = 150
    $txtProtoFilter.Height = 24
    $txtProtoFilter.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $lblProtoMax = New-InlineLabel -Text 'MaxFiles:' -MarginLeft 6
    $nudProtoMax = New-Object System.Windows.Forms.NumericUpDown
    $nudProtoMax.Width = 60
    $nudProtoMax.Minimum = 1
    $nudProtoMax.Maximum = 500
    $nudProtoMax.Value = 50
    $nudProtoMax.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnProtoLoad = New-Btn -Text 'Load & Parse' -W 100 -Color 'Blue'
    $protoToolbar.Controls.AddRange(@($lblProtoPath, $txtProtoPath, $btnProtoBrowse, $lblProtoFilter, $txtProtoFilter, $lblProtoMax, $nudProtoMax, $btnProtoLoad))

    $dgvProtocol = New-StyledDGV

    $protoBottom = New-FlowBar -H 38
    $protoBottom.Dock = 'Bottom'
    $btnProtoExport = New-Btn -Text 'Export...' -W 80
    $lblProtoCount = New-InlineLabel -Text '0 entries' -MarginLeft 10
    $protoBottom.Controls.AddRange(@($btnProtoExport, $lblProtoCount))

    $protoPanel.Controls.Add($dgvProtocol)
    $protoPanel.Controls.Add($protoBottom)
    $protoPanel.Controls.Add($protoToolbar)
    $tabProtocol.Controls.Add($protoPanel)

    $btnProtoBrowse.Add_Click({
        try {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($fbd.ShowDialog() -eq 'OK') { $txtProtoPath.Text = $fbd.SelectedPath }
        } catch {}
    })

    $btnProtoLoad.Add_Click({
        if (-not $txtProtoPath.Text) { return }
        Update-StatusBar 'Parsing protocol logs...'
        $logPath = $txtProtoPath.Text
        $maxFiles = [int]$nudProtoMax.Value
        $filterText = $txtProtoFilter.Text
        Start-AsyncJob -Name 'ProtocolLogs' -Form $form -ScriptBlock {
            param($LogPath, $MaxFiles, $Filter)
            $p = @{ LogPath = $LogPath; MaxFiles = $MaxFiles }
            if ($Filter) { $p['Filter'] = $Filter }
            return @(Parse-SmtpProtocolLog @p)
        } -Parameters @{ LogPath = $logPath; MaxFiles = $maxFiles; Filter = $filterText } -OnComplete {
            param($result)
            try {
                $script:LastProtocolData = @($result)
                Set-DGVData -DGV $dgvProtocol -Data $script:LastProtocolData
                $lblProtoCount.Text = "$($script:LastProtocolData.Count) entries"
                Update-StatusBar "Protocol logs: $($script:LastProtocolData.Count) entries"
            } catch {
                Update-StatusBar "Protocol UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Protocol log error: $err"
        }
    })
    $btnProtoExport.Add_Click({ Show-Export -Data $script:LastProtocolData -DefaultName 'protocol-logs' })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 5: LOG SEARCH
    # ═══════════════════════════════════════════════════════════════════════════
    $tabLogSearch = New-Object System.Windows.Forms.TabPage
    $tabLogSearch.Text = 'Log Search'

    $logSearchPanel = New-Object System.Windows.Forms.Panel
    $logSearchPanel.Dock = 'Fill'

    $logSearchToolbar = New-FlowBar -H 38
    $lblLogPath = New-InlineLabel -Text 'Path:'
    $txtLogPath = New-Object System.Windows.Forms.TextBox
    $txtLogPath.Width = 300
    $txtLogPath.Height = 24
    $txtLogPath.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnLogBrowse = New-Btn -Text 'Browse...' -W 80
    $lblLogPattern = New-InlineLabel -Text 'Pattern:' -MarginLeft 10
    $txtLogPattern = New-Object System.Windows.Forms.TextBox
    $txtLogPattern.Width = 200
    $txtLogPattern.Height = 24
    $txtLogPattern.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnLogSearch = New-Btn -Text 'Search' -W 80 -Color 'Blue'
    $logSearchToolbar.Controls.AddRange(@($lblLogPath, $txtLogPath, $btnLogBrowse, $lblLogPattern, $txtLogPattern, $btnLogSearch))

    $logSplit = New-Object System.Windows.Forms.SplitContainer
    $logSplit.Dock = 'Fill'
    $logSplit.Orientation = 'Horizontal'
    $logSplit.SplitterDistance = 300

    $dgvLogSearch = New-StyledDGV
    $logSplit.Panel1.Controls.Add($dgvLogSearch)

    $txtLogContext = New-ConsoleTextBox
    $logSplit.Panel2.Controls.Add($txtLogContext)

    $logSearchBottom = New-FlowBar -H 38
    $logSearchBottom.Dock = 'Bottom'
    $btnLogExport = New-Btn -Text 'Export...' -W 80
    $logSearchBottom.Controls.Add($btnLogExport)

    $logSearchPanel.Controls.Add($logSplit)
    $logSearchPanel.Controls.Add($logSearchBottom)
    $logSearchPanel.Controls.Add($logSearchToolbar)
    $tabLogSearch.Controls.Add($logSearchPanel)

    $btnLogBrowse.Add_Click({
        try {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($fbd.ShowDialog() -eq 'OK') { $txtLogPath.Text = $fbd.SelectedPath }
        } catch {}
    })

    $btnLogSearch.Add_Click({
        if (-not $txtLogPath.Text -or -not $txtLogPattern.Text) { return }
        Update-StatusBar 'Searching logs...'
        $logPathVal = $txtLogPath.Text
        $patternVal = $txtLogPattern.Text
        Start-AsyncJob -Name 'LogSearch' -Form $form -ScriptBlock {
            param($LogPath, $Pattern)
            return @(Search-TransportLogs -LogPath $LogPath -Pattern $Pattern)
        } -Parameters @{ LogPath = $logPathVal; Pattern = $patternVal } -OnComplete {
            param($result)
            try {
                $script:LastLogSearchResults = @($result)
                Set-DGVData -DGV $dgvLogSearch -Data $script:LastLogSearchResults
                Update-StatusBar "Log search: $($script:LastLogSearchResults.Count) results"
            } catch {
                Update-StatusBar "Log search UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Log search error: $err"
        }
    })

    $dgvLogSearch.Add_CellClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $contextCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Context') { $contextCol = $c.Index; break } }
            if ($null -ne $contextCol) {
                $txtLogContext.Text = "$($dgvLogSearch.Rows[$e.RowIndex].Cells[$contextCol].Value)"
            } else {
                $sb = [System.Text.StringBuilder]::new()
                foreach ($cell in $dgvLogSearch.Rows[$e.RowIndex].Cells) {
                    [void]$sb.AppendLine("$($cell.OwningColumn.Name): $($cell.Value)")
                }
                $txtLogContext.Text = $sb.ToString()
            }
        } catch {}
    })
    $btnLogExport.Add_Click({ Show-Export -Data $script:LastLogSearchResults -DefaultName 'log-search' })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 6: HEADER ANALYZER
    # ═══════════════════════════════════════════════════════════════════════════
    $tabHeaders = New-Object System.Windows.Forms.TabPage
    $tabHeaders.Text = 'Header Analyzer'

    $headerSplit = New-Object System.Windows.Forms.SplitContainer
    $headerSplit.Dock = 'Fill'
    $headerSplit.Orientation = 'Horizontal'
    $headerSplit.SplitterDistance = 200

    # Top: input area
    $headerInputPanel = New-Object System.Windows.Forms.Panel
    $headerInputPanel.Dock = 'Fill'

    $headerToolbar = New-FlowBar -H 38
    $btnHeaderAnalyze = New-Btn -Text 'Analyze' -W 80 -Color 'Blue'
    $btnHeaderLoad = New-Btn -Text 'Load File...' -W 100
    $btnHeaderExport = New-Btn -Text 'Export...' -W 80
    $headerToolbar.Controls.AddRange(@($btnHeaderAnalyze, $btnHeaderLoad, $btnHeaderExport))

    $txtHeaderInput = New-Object System.Windows.Forms.TextBox
    $txtHeaderInput.Multiline = $true
    $txtHeaderInput.ScrollBars = 'Both'
    $txtHeaderInput.WordWrap = $false
    $txtHeaderInput.Dock = 'Fill'
    $txtHeaderInput.Font = New-Object System.Drawing.Font('Consolas', 10)

    $headerInputPanel.Controls.Add($txtHeaderInput)
    $headerInputPanel.Controls.Add($headerToolbar)
    $headerSplit.Panel1.Controls.Add($headerInputPanel)

    # Bottom: results (split vertical: left=info+hops, right=x-headers)
    $headerResultSplit = New-Object System.Windows.Forms.SplitContainer
    $headerResultSplit.Dock = 'Fill'
    $headerResultSplit.Orientation = 'Vertical'
    $headerResultSplit.SplitterDistance = 550

    # Left side: info + hops
    $headerLeftPanel = New-Object System.Windows.Forms.Panel
    $headerLeftPanel.Dock = 'Fill'

    $txtHeaderInfo = New-Object System.Windows.Forms.TextBox
    $txtHeaderInfo.Multiline = $true
    $txtHeaderInfo.ReadOnly = $true
    $txtHeaderInfo.ScrollBars = 'Vertical'
    $txtHeaderInfo.Dock = 'Top'
    $txtHeaderInfo.Height = 160
    $txtHeaderInfo.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtHeaderInfo.BackColor = [System.Drawing.Color]::FromArgb(250,250,255)

    $dgvHops = New-StyledDGV
    $dgvHops.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $delayCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Delay') { $delayCol = $c.Index; break } }
            if ($null -ne $delayCol) {
                $val = "$($row.Cells[$delayCol].Value)"
                $sec = 0
                if ($val -match '(\d+(\.\d+)?)') { $sec = [double]$Matches[1] }
                if ($sec -gt 5) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,180,180) }
                elseif ($sec -gt 1) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,180) }
                else { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
            }
        } catch {}
    })

    $headerLeftPanel.Controls.Add($dgvHops)
    $headerLeftPanel.Controls.Add($txtHeaderInfo)
    $headerResultSplit.Panel1.Controls.Add($headerLeftPanel)

    # Right side: X-Headers
    $dgvXHeaders = New-StyledDGV
    $headerResultSplit.Panel2.Controls.Add($dgvXHeaders)

    $headerSplit.Panel2.Controls.Add($headerResultSplit)
    $tabHeaders.Controls.Add($headerSplit)

    $btnHeaderLoad.Add_Click({
        try {
            $ofd = New-Object System.Windows.Forms.OpenFileDialog
            $ofd.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
            if ($ofd.ShowDialog() -eq 'OK') {
                $txtHeaderInput.Text = [System.IO.File]::ReadAllText($ofd.FileName)
            }
        } catch { Update-StatusBar "File load error: $_" }
    })

    $btnHeaderAnalyze.Add_Click({
        if (-not $txtHeaderInput.Text) { return }
        Update-StatusBar 'Analyzing headers...'
        try {
            # Header parsing is CPU-only, no Exchange calls, runs fast — keep sync
            $result = Parse-EmailHeaders -HeaderText $txtHeaderInput.Text
            $script:LastHeaderResult = $result

            # Info textbox
            $sb = [System.Text.StringBuilder]::new()
            if ($result.PSObject.Properties['MessageId'])  { [void]$sb.AppendLine("Message-ID : $($result.MessageId)") }
            if ($result.PSObject.Properties['From'])       { [void]$sb.AppendLine("From       : $($result.From)") }
            if ($result.PSObject.Properties['To'])         { [void]$sb.AppendLine("To         : $($result.To)") }
            if ($result.PSObject.Properties['Subject'])    { [void]$sb.AppendLine("Subject    : $($result.Subject)") }
            if ($result.PSObject.Properties['Date'])       { [void]$sb.AppendLine("Date       : $($result.Date)") }
            if ($result.PSObject.Properties['ReturnPath']) { [void]$sb.AppendLine("Return-Path: $($result.ReturnPath)") }
            if ($result.PSObject.Properties['SPF'])        { [void]$sb.AppendLine("SPF        : $($result.SPF)") }
            if ($result.PSObject.Properties['DKIM'])       { [void]$sb.AppendLine("DKIM       : $($result.DKIM)") }
            if ($result.PSObject.Properties['DMARC'])      { [void]$sb.AppendLine("DMARC      : $($result.DMARC)") }
            if ($result.PSObject.Properties['Hops']) {
                $hopCount = ($result.Hops | Measure-Object).Count
                [void]$sb.AppendLine("Total Hops : $hopCount")
            }
            if ($result.PSObject.Properties['TotalDelay']) { [void]$sb.AppendLine("Total Delay: $($result.TotalDelay)") }
            $txtHeaderInfo.Text = $sb.ToString()

            # Hops grid
            if ($result.PSObject.Properties['Hops'] -and $result.Hops) {
                $hopsForDisplay = @()
                $hopNum = 1
                foreach ($hop in $result.Hops) {
                    $hopsForDisplay += [PSCustomObject]@{
                        '#'         = $hopNum
                        From        = "$($hop.From)"
                        By          = "$($hop.By)"
                        Protocol    = "$($hop.Protocol)"
                        TLS         = "$($hop.TLS)"
                        Timestamp   = "$($hop.Timestamp)"
                        Delay       = "$($hop.Delay)"
                    }
                    $hopNum++
                }
                Set-DGVData -DGV $dgvHops -Data $hopsForDisplay
            }

            # X-Headers
            if ($result.PSObject.Properties['XHeaders'] -and $result.XHeaders) {
                $xhData = @()
                foreach ($xh in $result.XHeaders) {
                    $xhData += [PSCustomObject]@{
                        Header = "$($xh.Name)"
                        Value  = "$($xh.Value)"
                    }
                }
                Set-DGVData -DGV $dgvXHeaders -Data $xhData
            }

            Update-StatusBar 'Header analysis complete'
        } catch { Update-StatusBar "Header error: $_" }
    })

    $btnHeaderExport.Add_Click({
        try {
            if ($script:LastHeaderResult -and $script:LastHeaderResult.PSObject.Properties['Hops']) {
                Show-Export -Data @($script:LastHeaderResult.Hops) -DefaultName 'header-hops'
            }
        } catch {}
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 7: DIAGNOSTICS (inner TabControl)
    # ═══════════════════════════════════════════════════════════════════════════
    $tabDiag = New-Object System.Windows.Forms.TabPage
    $tabDiag.Text = 'Diagnostics'

    $diagTabs = New-Object System.Windows.Forms.TabControl
    $diagTabs.Dock = 'Fill'

    # --- DNS sub-tab ---
    $tabDNS = New-Object System.Windows.Forms.TabPage
    $tabDNS.Text = 'DNS'

    $dnsToolbar = New-FlowBar -H 38
    $lblDnsDomain = New-InlineLabel -Text 'Domain:'
    $txtDnsDomain = New-Object System.Windows.Forms.TextBox
    $txtDnsDomain.Width = 250
    $txtDnsDomain.Height = 24
    $txtDnsDomain.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnDnsCheck = New-Btn -Text 'Check' -W 80 -Color 'Blue'
    $dnsToolbar.Controls.AddRange(@($lblDnsDomain, $txtDnsDomain, $btnDnsCheck))

    $txtDnsResults = New-ConsoleTextBox
    $tabDNS.Controls.Add($txtDnsResults)
    $tabDNS.Controls.Add($dnsToolbar)

    $btnDnsCheck.Add_Click({
        if (-not $txtDnsDomain.Text) { return }
        $domain = $txtDnsDomain.Text
        Update-StatusBar "Checking DNS for $domain..."
        Start-AsyncJob -Name "DNS $domain" -Form $form -ScriptBlock {
            param($Domain)
            return Test-DomainMailHealth -Domain $Domain
        } -Parameters @{ Domain = $domain } -OnComplete {
            param($result)
            try {
                $sb = [System.Text.StringBuilder]::new()
                [void]$sb.AppendLine("=" * 60)
                [void]$sb.AppendLine("DOMAIN MAIL HEALTH: $($result.Domain)")
                [void]$sb.AppendLine("=" * 60)

                $healthColor = switch ($result.Health) { 'Good' { 'GOOD' } 'Warning' { 'WARNING' } 'Bad' { 'BAD' } default { $result.Health } }
                [void]$sb.AppendLine("Health: $healthColor")
                if ($result.Issues.Count -gt 0) {
                    [void]$sb.AppendLine("Issues: $($result.Issues -join '; ')")
                }
                [void]$sb.AppendLine("")

                [void]$sb.AppendLine("--- MX RECORDS ---")
                if ($result.MX -and ($result.MX | Measure-Object).Count -gt 0) {
                    foreach ($mx in $result.MX) {
                        [void]$sb.AppendLine("  Priority: $($mx.Priority)  Server: $($mx.MailServer)  IP: $($mx.IP)")
                    }
                } else { [void]$sb.AppendLine("  No MX records found") }
                [void]$sb.AppendLine("")

                [void]$sb.AppendLine("--- SPF ---")
                if ($result.SPF.Record) {
                    [void]$sb.AppendLine("  Record: $($result.SPF.Record)")
                    [void]$sb.AppendLine("  Qualifier: $($result.SPF.Qualifier)")
                    if ($result.SPF.Mechanisms) {
                        [void]$sb.AppendLine("  Mechanisms:")
                        foreach ($m in $result.SPF.Mechanisms) {
                            [void]$sb.AppendLine("    $($m.Qualifier)$($m.Mechanism)")
                        }
                    }
                } else { [void]$sb.AppendLine("  No SPF record") }
                [void]$sb.AppendLine("")

                [void]$sb.AppendLine("--- DKIM ---")
                if ($result.DKIM.Record) {
                    [void]$sb.AppendLine("  Selector: $($result.DKIM.Selector)")
                    [void]$sb.AppendLine("  Key Type: $($result.DKIM.KeyType)")
                    [void]$sb.AppendLine("  Record: $($result.DKIM.Record)")
                } else { [void]$sb.AppendLine("  No DKIM record found") }
                [void]$sb.AppendLine("")

                [void]$sb.AppendLine("--- DMARC ---")
                if ($result.DMARC.Record) {
                    [void]$sb.AppendLine("  Record: $($result.DMARC.Record)")
                    [void]$sb.AppendLine("  Policy: $($result.DMARC.Policy)")
                    [void]$sb.AppendLine("  Subdomain Policy: $($result.DMARC.SubdomainPolicy)")
                    [void]$sb.AppendLine("  Percentage: $($result.DMARC.Percentage)")
                    [void]$sb.AppendLine("  Report URI: $($result.DMARC.ReportUri)")
                    [void]$sb.AppendLine("  DKIM Alignment: $($result.DMARC.DkimAlignment)")
                    [void]$sb.AppendLine("  SPF Alignment: $($result.DMARC.SpfAlignment)")
                } else { [void]$sb.AppendLine("  No DMARC record") }

                $txtDnsResults.Text = $sb.ToString()
                Update-StatusBar "DNS check complete: $healthColor"
            } catch {
                Update-StatusBar "DNS UI error: $_"
            }
        } -OnError {
            param($err)
            $txtDnsResults.Text = "Error: $err"
            Update-StatusBar "DNS error: $err"
        }
    })

    # --- Transport Rules sub-tab ---
    $tabRules = New-Object System.Windows.Forms.TabPage
    $tabRules.Text = 'Transport Rules'

    $rulesToolbar = New-FlowBar -H 38
    $btnLoadRules = New-Btn -Text 'Load Rules' -W 100 -Color 'Blue'
    $rulesToolbar.Controls.Add($btnLoadRules)

    $dgvRules = New-StyledDGV
    $tabRules.Controls.Add($dgvRules)
    $tabRules.Controls.Add($rulesToolbar)

    $btnLoadRules.Add_Click({
        Update-StatusBar 'Loading transport rules...'
        Start-AsyncJob -Name 'TransportRules' -Form $form -ScriptBlock {
            return @(Get-ExchangeTransportRules)
        } -OnComplete {
            param($result)
            try {
                $script:LastRulesData = @($result)
                $display = $result | ForEach-Object {
                    [PSCustomObject]@{
                        Name       = $_.Name
                        State      = $_.State
                        Priority   = $_.Priority
                        Mode       = $_.Mode
                        Conditions = $_.Conditions
                        Actions    = $_.Actions
                    }
                }
                Set-DGVData -DGV $dgvRules -Data @($display)
                Update-StatusBar "Transport rules: $($script:LastRulesData.Count) loaded"
            } catch {
                Update-StatusBar "Rules UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Rules error: $err"
        }
    })

    # Double-click rule -> details popup
    $dgvRules.Add_CellDoubleClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $ruleName = "$($dgvRules.Rows[$e.RowIndex].Cells['Name'].Value)"
            $ruleObj = $script:LastRulesData | Where-Object { $_.Name -eq $ruleName } | Select-Object -First 1
            if (-not $ruleObj) { return }
            $detailForm = New-Object System.Windows.Forms.Form
            $detailForm.Text = "Rule: $ruleName"
            $detailForm.Size = New-Object System.Drawing.Size(600, 400)
            $detailForm.StartPosition = 'CenterParent'
            $detailTxt = New-ConsoleTextBox
            $sb = [System.Text.StringBuilder]::new()
            foreach ($prop in $ruleObj.PSObject.Properties) {
                [void]$sb.AppendLine("$($prop.Name): $($prop.Value)")
            }
            $detailTxt.Text = $sb.ToString()
            $detailForm.Controls.Add($detailTxt)
            $detailForm.ShowDialog()
        } catch {}
    })

    # --- Certificates sub-tab ---
    $tabCerts = New-Object System.Windows.Forms.TabPage
    $tabCerts.Text = 'Certificates'

    $certsToolbar = New-FlowBar -H 38
    $btnLoadCerts = New-Btn -Text 'Load Certificates' -W 130 -Color 'Blue'
    $certsToolbar.Controls.Add($btnLoadCerts)

    $certSplit = New-Object System.Windows.Forms.SplitContainer
    $certSplit.Dock = 'Fill'
    $certSplit.Orientation = 'Horizontal'
    $certSplit.SplitterDistance = 300

    $dgvCerts = New-StyledDGV
    $dgvCerts.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $row = $s.Rows[$e.RowIndex]
            $statusCol = $null
            foreach ($c in $s.Columns) { if ($c.Name -eq 'Status') { $statusCol = $c.Index; break } }
            if ($null -ne $statusCol) {
                $val = "$($row.Cells[$statusCol].Value)"
                if ($val -eq 'Expired')  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,180,180) }
                elseif ($val -eq 'Expiring') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,255,180) }
                elseif ($val -eq 'Valid')    { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(200,255,200) }
            }
        } catch {}
    })
    $certSplit.Panel1.Controls.Add($dgvCerts)

    $txtCertBindings = New-ConsoleTextBox
    $certSplit.Panel2.Controls.Add($txtCertBindings)

    $tabCerts.Controls.Add($certSplit)
    $tabCerts.Controls.Add($certsToolbar)

    $btnLoadCerts.Add_Click({
        Update-StatusBar 'Loading certificates...'
        Start-AsyncJob -Name 'Certificates' -Form $form -ScriptBlock {
            $certs = @(Get-ExchangeCertificates)
            $bindings = @(Get-ConnectorCertificateBindings)
            return @{ Certs = $certs; Bindings = $bindings }
        } -OnComplete {
            param($result)
            try {
                $certs = $result.Certs
                $bindings = $result.Bindings
                $script:LastCertData = @($certs)
                $display = $certs | ForEach-Object {
                    [PSCustomObject]@{
                        Subject         = $_.Subject
                        Issuer          = $_.Issuer
                        NotAfter        = "$($_.NotAfter)"
                        DaysUntilExpiry = $_.DaysUntilExpiry
                        Status          = $_.Status
                        Services        = "$($_.Services)"
                    }
                }
                Set-DGVData -DGV $dgvCerts -Data @($display)

                $sb = [System.Text.StringBuilder]::new()
                [void]$sb.AppendLine("CONNECTOR CERTIFICATE BINDINGS")
                [void]$sb.AppendLine("=" * 60)
                foreach ($b in $bindings) {
                    [void]$sb.AppendLine("[$($b.ConnectorType)] $($b.ConnectorName)")
                    if ($b.CertSubject) {
                        [void]$sb.AppendLine("  Cert: $($b.CertSubject)  Status: $($b.CertStatus)  Expires: $($b.CertExpiry)")
                    } else {
                        [void]$sb.AppendLine("  No certificate bound")
                    }
                    [void]$sb.AppendLine("")
                }
                $txtCertBindings.Text = $sb.ToString()
                Update-StatusBar "Certificates: $($script:LastCertData.Count) loaded"
            } catch {
                Update-StatusBar "Cert UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Cert error: $err"
        }
    })

    # --- Connectivity Logs sub-tab ---
    $tabConnLogs = New-Object System.Windows.Forms.TabPage
    $tabConnLogs.Text = 'Connectivity Logs'

    $connLogToolbar = New-FlowBar -H 38
    $lblConnLogPath = New-InlineLabel -Text 'Log Path:'
    $txtConnLogPath = New-Object System.Windows.Forms.TextBox
    $txtConnLogPath.Width = 300
    $txtConnLogPath.Height = 24
    $txtConnLogPath.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    $btnConnLogBrowse = New-Btn -Text 'Browse...' -W 80
    $btnConnLogLoad = New-Btn -Text 'Load' -W 70 -Color 'Blue'
    $connLogToolbar.Controls.AddRange(@($lblConnLogPath, $txtConnLogPath, $btnConnLogBrowse, $btnConnLogLoad))

    $dgvConnLog = New-StyledDGV
    $tabConnLogs.Controls.Add($dgvConnLog)
    $tabConnLogs.Controls.Add($connLogToolbar)

    $btnConnLogBrowse.Add_Click({
        try {
            $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($fbd.ShowDialog() -eq 'OK') { $txtConnLogPath.Text = $fbd.SelectedPath }
        } catch {}
    })

    $btnConnLogLoad.Add_Click({
        if (-not $txtConnLogPath.Text) { return }
        $connPath = $txtConnLogPath.Text
        Update-StatusBar 'Parsing connectivity logs...'
        Start-AsyncJob -Name 'ConnectivityLogs' -Form $form -ScriptBlock {
            param($LogPath)
            return @(Parse-ConnectivityLog -LogPath $LogPath)
        } -Parameters @{ LogPath = $connPath } -OnComplete {
            param($result)
            try {
                Set-DGVData -DGV $dgvConnLog -Data @($result)
                Update-StatusBar "Connectivity logs: $($result.Count) entries"
            } catch {
                Update-StatusBar "Connectivity log UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Connectivity log error: $err"
        }
    })

    $diagTabs.TabPages.AddRange(@($tabDNS, $tabRules, $tabCerts, $tabConnLogs))
    $tabDiag.Controls.Add($diagTabs)

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 8: STATISTICS
    # ═══════════════════════════════════════════════════════════════════════════
    $tabStats = New-Object System.Windows.Forms.TabPage
    $tabStats.Text = 'Statistics'

    $statsPanel = New-Object System.Windows.Forms.Panel
    $statsPanel.Dock = 'Fill'

    $statsToolbar = New-FlowBar -H 38
    $btnLoadStats = New-Btn -Text 'Load Statistics' -W 120 -Color 'Blue'
    $lblStatsPeriod = New-InlineLabel -Text 'Period:' -MarginLeft 10
    $cmbStatsPeriod = New-Object System.Windows.Forms.ComboBox
    $cmbStatsPeriod.DropDownStyle = 'DropDownList'
    $cmbStatsPeriod.Width = 100
    $cmbStatsPeriod.Margin = New-Object System.Windows.Forms.Padding(3,6,3,4)
    @('Last 1h','Last 6h','Last 24h','Last 7d') | ForEach-Object { [void]$cmbStatsPeriod.Items.Add($_) }
    $cmbStatsPeriod.SelectedIndex = 2
    $btnStatsExport = New-Btn -Text 'Export...' -W 80
    $statsToolbar.Controls.AddRange(@($btnLoadStats, $lblStatsPeriod, $cmbStatsPeriod, $btnStatsExport))

    $statsTabs = New-Object System.Windows.Forms.TabControl
    $statsTabs.Dock = 'Fill'

    $tabStatSender = New-Object System.Windows.Forms.TabPage
    $tabStatSender.Text = 'By Sender'
    $dgvStatSender = New-StyledDGV
    $tabStatSender.Controls.Add($dgvStatSender)

    $tabStatRecip = New-Object System.Windows.Forms.TabPage
    $tabStatRecip.Text = 'By Recipient'
    $dgvStatRecip = New-StyledDGV
    $tabStatRecip.Controls.Add($dgvStatRecip)

    $tabStatDomain = New-Object System.Windows.Forms.TabPage
    $tabStatDomain.Text = 'By Domain'
    $dgvStatDomain = New-StyledDGV
    $tabStatDomain.Controls.Add($dgvStatDomain)

    $tabStatHourly = New-Object System.Windows.Forms.TabPage
    $tabStatHourly.Text = 'By Hour'
    $dgvStatHourly = New-StyledDGV
    $tabStatHourly.Controls.Add($dgvStatHourly)

    $tabStatConnector = New-Object System.Windows.Forms.TabPage
    $tabStatConnector.Text = 'By Connector'
    $dgvStatConnector = New-StyledDGV
    $tabStatConnector.Controls.Add($dgvStatConnector)

    $statsTabs.TabPages.AddRange(@($tabStatSender, $tabStatRecip, $tabStatDomain, $tabStatHourly, $tabStatConnector))

    $statsPanel.Controls.Add($statsTabs)
    $statsPanel.Controls.Add($statsToolbar)
    $tabStats.Controls.Add($statsPanel)

    $script:StatsByTab = @{
        Sender    = @()
        Recipient = @()
        Domain    = @()
        Hourly    = @()
        Connector = @()
    }

    $btnLoadStats.Add_Click({
        Update-StatusBar 'Loading statistics...'
        $periodText = $cmbStatsPeriod.SelectedItem.ToString()
        $hours = switch ($periodText) {
            'Last 1h'  { 1 }
            'Last 6h'  { 6 }
            'Last 24h' { 24 }
            'Last 7d'  { 168 }
            default    { 24 }
        }
        $statsParams = @{
            Start = (Get-Date).AddHours(-$hours)
            End   = Get-Date
        }
        $scopeVal = Get-SelectedScope
        if ($scopeVal) { $statsParams['Server'] = $scopeVal }

        Start-AsyncJob -Name "Statistics ($periodText)" -Form $form -ScriptBlock {
            param($Params)
            $trackData = Trace-ExchangeMessage @Params
            return @{
                Sender    = @(Get-SenderStatistics -TrackingData $trackData)
                Recipient = @(Get-RecipientStatistics -TrackingData $trackData)
                Domain    = @(Get-DomainStatistics -TrackingData $trackData)
                Hourly    = @(Get-HourlyStatistics -TrackingData $trackData)
                Connector = @(Get-ConnectorStatistics -TrackingData $trackData)
            }
        } -Parameters @{ Params = $statsParams } -OnComplete {
            param($result)
            try {
                $script:StatsByTab.Sender    = @($result.Sender)
                $script:StatsByTab.Recipient = @($result.Recipient)
                $script:StatsByTab.Domain    = @($result.Domain)
                $script:StatsByTab.Hourly    = @($result.Hourly)
                $script:StatsByTab.Connector = @($result.Connector)

                Set-DGVData -DGV $dgvStatSender    -Data $script:StatsByTab.Sender
                Set-DGVData -DGV $dgvStatRecip     -Data $script:StatsByTab.Recipient
                Set-DGVData -DGV $dgvStatDomain    -Data $script:StatsByTab.Domain
                Set-DGVData -DGV $dgvStatHourly    -Data $script:StatsByTab.Hourly
                Set-DGVData -DGV $dgvStatConnector -Data $script:StatsByTab.Connector

                Update-StatusBar "Statistics loaded ($periodText)"
            } catch {
                Update-StatusBar "Statistics UI error: $_"
            }
        } -OnError {
            param($err)
            Update-StatusBar "Statistics error: $err"
        }
    })

    $btnStatsExport.Add_Click({
        try {
            $currentSubTab = $statsTabs.SelectedTab.Text
            $data = switch ($currentSubTab) {
                'By Sender'    { $script:StatsByTab.Sender }
                'By Recipient' { $script:StatsByTab.Recipient }
                'By Domain'    { $script:StatsByTab.Domain }
                'By Hour'      { $script:StatsByTab.Hourly }
                'By Connector' { $script:StatsByTab.Connector }
                default        { @() }
            }
            Show-Export -Data $data -DefaultName "stats-$($currentSubTab -replace ' ','-')"
        } catch {}
    })

    # ═══════════════════════════════════════════════════════════════════════════
    # TAB 9: REPORTS
    # ═══════════════════════════════════════════════════════════════════════════
    $tabReports = New-Object System.Windows.Forms.TabPage
    $tabReports.Text = 'Reports'

    $reportsPanel = New-Object System.Windows.Forms.Panel
    $reportsPanel.Dock = 'Fill'

    $reportsToolbar = New-FlowBar -H 38
    $lblReportType = New-InlineLabel -Text 'Report:'
    $cmbReportType = New-Object System.Windows.Forms.ComboBox
    $cmbReportType.DropDownStyle = 'DropDownList'
    $cmbReportType.Width = 160
    $cmbReportType.Margin = New-Object System.Windows.Forms.Padding(3,6,10,4)
    @('Full','Queues','Connectors','AgentLog','RoutingTable','DSN','Summary','Pipeline','BackPressure') | ForEach-Object { [void]$cmbReportType.Items.Add($_) }
    $cmbReportType.SelectedIndex = 0
    $btnRunReport = New-Btn -Text 'Run Report' -W 100 -Color 'Blue'
    $btnSaveReport = New-Btn -Text 'Save to File...' -W 110
    $reportsToolbar.Controls.AddRange(@($lblReportType, $cmbReportType, $btnRunReport, $btnSaveReport))

    $txtReport = New-Object System.Windows.Forms.TextBox
    $txtReport.Multiline = $true
    $txtReport.ScrollBars = 'Both'
    $txtReport.WordWrap = $false
    $txtReport.ReadOnly = $true
    $txtReport.Font = New-Object System.Drawing.Font('Consolas', 10)
    $txtReport.BackColor = [System.Drawing.Color]::FromArgb(25,25,35)
    $txtReport.ForeColor = [System.Drawing.Color]::FromArgb(220,220,220)
    $txtReport.Dock = 'Fill'

    $reportsPanel.Controls.Add($txtReport)
    $reportsPanel.Controls.Add($reportsToolbar)
    $tabReports.Controls.Add($reportsPanel)

    $btnRunReport.Add_Click({
        $reportType = $cmbReportType.SelectedItem.ToString()
        Update-StatusBar "Running $reportType report..."
        $scopeVal = Get-SelectedScope
        Start-AsyncJob -Name "Report: $reportType" -Form $form -ScriptBlock {
            param($ReportType, $Server)
            $p = @{ ReportType = $ReportType }
            if ($Server) { $p['Server'] = $Server }
            return (Get-TransportReportData @p)
        } -Parameters @{ ReportType = $reportType; Server = $scopeVal } -OnComplete {
            param($result)
            try {
                $script:LastReportText = $result
                $txtReport.Text = ($result | Out-String)
                try { Write-OperatorLog -Action 'RunReport' -Target $reportType } catch {}
                Update-StatusBar "$reportType report generated"
            } catch {
                Update-StatusBar "Report UI error: $_"
            }
        } -OnError {
            param($err)
            $txtReport.Text = "Error: $err"
            Update-StatusBar "Report error: $err"
        }
    })

    $btnSaveReport.Add_Click({
        try {
            if (-not $script:LastReportText) {
                [System.Windows.Forms.MessageBox]::Show('No report to save.','Save','OK','Information')
                return
            }
            $sfd = New-Object System.Windows.Forms.SaveFileDialog
            $sfd.Filter = 'Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
            $sfd.FileName = "report-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
            if ($sfd.ShowDialog() -eq 'OK') {
                ($script:LastReportText | Out-String) | Set-Content -Path $sfd.FileName -Encoding UTF8
                Update-StatusBar "Report saved to $($sfd.FileName)"
            }
        } catch { Update-StatusBar "Save error: $_" }
    })

    # ─── Assemble tabs ───────────────────────────────────────────────────────
    $tabs.TabPages.AddRange(@($tabDash, $tabQueues, $tabTracking, $tabProtocol, $tabLogSearch, $tabHeaders, $tabDiag, $tabStats, $tabReports))
    $form.Controls.Add($tabs)

    # ─── Connect button (async) ──────────────────────────────────────────────
    $btnConnect.Add_Click({
        $server = $txtServer.Text.Trim()
        if (-not $server) {
            [System.Windows.Forms.MessageBox]::Show('Enter an Exchange server name.','Connect','OK','Warning')
            return
        }
        Update-StatusBar "Connecting to $server..."
        $lblConnStatus.Text = 'Connecting...'
        $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(200,150,0)
        $btnConnect.Enabled = $false

        Start-AsyncJob -Name "Connect $server" -Form $form -ScriptBlock {
            param($Server)
            $session = Connect-ExchangeRemote -Server $Server
            $transportServers = @()
            try { $transportServers = @(Get-ExchangeTransportServers) } catch {}
            return @{ Session = $session; TransportServers = $transportServers }
        } -Parameters @{ Server = $server } -OnComplete {
            param($result)
            try {
                # Disconnect old session if exists
                if ($script:Session) {
                    try { Disconnect-ExchangeRemote -Session $script:Session } catch {}
                }
                $script:Session = $result.Session
                $script:TransportServers = $result.TransportServers

                $lblConnStatus.Text = "Connected: $server"
                $lblConnStatus.ForeColor = [System.Drawing.Color]::Green
                $btnConnect.Enabled = $true
                $btnDisconnect.Visible = $true
                try { Update-RecentServers -Server $server } catch {}

                # Populate scope combo
                $script:ScopeCombo.Items.Clear()
                [void]$script:ScopeCombo.Items.Add('(All Servers)')
                foreach ($ts in $script:TransportServers) {
                    [void]$script:ScopeCombo.Items.Add($ts.Name)
                }
                $script:ScopeCombo.SelectedIndex = 0

                try { Write-OperatorLog -Action 'Connect' -Target $server } catch {}
                Update-StatusBar "Connected to $server"

                # Auto-refresh initial data
                & $refreshDashboard
                & $refreshQueues
                & $refreshErrors
            } catch {
                $lblConnStatus.Text = 'Connection setup error'
                $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
                $btnConnect.Enabled = $true
                Update-StatusBar "Connection setup error: $_"
            }
        } -OnError {
            param($err)
            $lblConnStatus.Text = 'Connection failed'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Red
            $btnConnect.Enabled = $true
            Update-StatusBar "Connection failed: $err"
            [System.Windows.Forms.MessageBox]::Show("Failed to connect: $err",'Connection Error','OK','Error')
        }
    })

    $btnDisconnect.Add_Click({
        try {
            if ($script:Session) {
                try { Disconnect-ExchangeRemote -Session $script:Session } catch {}
                $script:Session = $null
            }
            $lblConnStatus.Text = 'Disconnected'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::Gray
            $btnDisconnect.Visible = $false
            Update-StatusBar 'Disconnected'
        } catch {}
    })

    # ─── Async Poller Timer ──────────────────────────────────────────────────
    $asyncPoller = New-AsyncPollerTimer

    # ─── Global auto-refresh timer ───────────────────────────────────────────
    $autoRefreshTimer = New-Object System.Windows.Forms.Timer
    $interval = try { if ($script:Settings.RefreshIntervalSec) { [int]$script:Settings.RefreshIntervalSec * 1000 } else { 30000 } } catch { 30000 }
    $autoRefreshTimer.Interval = $interval

    $autoRefreshTimer.Add_Tick({
        if (-not $script:Session) { return }
        # Only auto-refresh if no jobs are currently running (avoid pileup)
        $running = Get-RunningJobCount
        if ($running -gt 0) { return }
        try {
            & $refreshDashboard
            & $refreshQueues
            & $refreshErrors
        } catch {}
    })

    $chkAutoRefresh.Add_CheckedChanged({
        if ($chkAutoRefresh.Checked) {
            $autoRefreshTimer.Start()
            Update-StatusBar 'Auto-refresh enabled'
        } else {
            $autoRefreshTimer.Stop()
            Update-StatusBar 'Auto-refresh disabled'
        }
    })

    # ─── Keyboard shortcuts (form-level KeyPreview) ──────────────────────────
    $form.Add_KeyDown({
        param($s, $e)
        try {
            # F5 = Refresh current tab
            if ($e.KeyCode -eq 'F5') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabDash     { & $refreshDashboard }
                    $tabQueues   { & $refreshQueues; & $refreshErrors }
                    $tabTracking { & $doTrackSearch }
                    default      { Update-StatusBar 'F5: No refresh action for this tab' }
                }
            }
            # Ctrl+E = Export
            if ($e.Control -and $e.KeyCode -eq 'E') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabQueues    { Show-Export -Data $script:LastQueueData -DefaultName 'queues' }
                    $tabTracking  { Show-Export -Data $script:LastTrackingResults -DefaultName 'tracking' }
                    $tabProtocol  { Show-Export -Data $script:LastProtocolData -DefaultName 'protocol' }
                    $tabLogSearch { Show-Export -Data $script:LastLogSearchResults -DefaultName 'log-search' }
                    default       { Update-StatusBar 'Ctrl+E: No export for this tab' }
                }
            }
            # Ctrl+F = Focus search/filter
            if ($e.Control -and $e.KeyCode -eq 'F') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabQueues    { $txtQueueFilter.Focus() }
                    $tabTracking  { $txtTrackMsgId.Focus() }
                    $tabProtocol  { $txtProtoFilter.Focus() }
                    $tabLogSearch { $txtLogPattern.Focus() }
                    default       {}
                }
            }
            # Escape = Clear filter/search
            if ($e.KeyCode -eq 'Escape') {
                $e.Handled = $true
                switch ($tabs.SelectedTab) {
                    $tabQueues    { $txtQueueFilter.Text = '' }
                    $tabTracking  { $txtTrackMsgId.Text = ''; $txtTrackSender.Text = ''; $txtTrackRecip.Text = ''; $txtTrackSubject.Text = '' }
                    $tabProtocol  { $txtProtoFilter.Text = '' }
                    $tabLogSearch { $txtLogPattern.Text = '' }
                    default       {}
                }
            }
        } catch {}
    })

    # ─── Form events ────────────────────────────────────────────────────────
    $form.Add_FormClosing({
        # Save settings
        try {
            $settings = Get-AppSettings
            $settings.WindowWidth = $form.Width
            $settings.WindowHeight = $form.Height
            $settings.LastServer = $txtServer.Text
            if ($script:ScopeCombo.SelectedItem) {
                $settings.LastScope = $script:ScopeCombo.SelectedItem.ToString()
            }
            $settings.SplitterPositions = @{
                Main   = $queueSplitOuter.SplitterDistance
                Detail = $queueSplitInner.SplitterDistance
            }
            Save-AppSettings -Settings $settings
        } catch {}

        # Stop timers
        try { $asyncPoller.Stop(); $asyncPoller.Dispose() } catch {}
        try { $autoRefreshTimer.Stop(); $autoRefreshTimer.Dispose() } catch {}

        # Cleanup async jobs
        try {
            foreach ($job in $script:AsyncJobs) {
                if ($job.Status -eq 'Running') {
                    try { $job.PowerShell.Stop() } catch {}
                    try { $job.PowerShell.Dispose() } catch {}
                    try { $job.Runspace.Close(); $job.Runspace.Dispose() } catch {}
                }
            }
        } catch {}

        # Disconnect session
        if ($script:Session) {
            try { Disconnect-ExchangeRemote -Session $script:Session } catch {}
        }
    })

    # ─── Show form ───────────────────────────────────────────────────────────
    [void]$form.ShowDialog()
}

# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

if ($MyInvocation.InvocationName -ne '.') {
    Show-ExchangeRetryGUI
}
