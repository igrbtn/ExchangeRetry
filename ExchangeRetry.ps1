<#
.SYNOPSIS
    ExchangeRetry — GUI-инструмент для управления очередями Microsoft Exchange.
.DESCRIPTION
    Позволяет просматривать очереди Exchange, фильтровать застрявшие письма
    и выполнять retry/suspend/remove для выбранных сообщений.
.NOTES
    Version: 0.1.0
    Requires: Exchange Management Shell (EMS) или удалённое подключение к Exchange.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─── Configuration ───────────────────────────────────────────────────────────

$script:Config = @{
    ExchangeServer  = $env:EXCHANGE_SERVER
    UseRemoteShell  = $true
    DefaultPageSize  = 100
    RefreshIntervalSec = 30
}

# ─── Exchange Connection ─────────────────────────────────────────────────────

function Connect-ExchangeRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Server,
        [PSCredential]$Credential
    )

    $uri = "http://$Server/PowerShell/"
    $params = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $uri
        Authentication    = 'Kerberos'
    }
    if ($Credential) {
        $params['Credential'] = $Credential
    }

    try {
        $session = New-PSSession @params -ErrorAction Stop
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        return $session
    }
    catch {
        throw "Failed to connect to Exchange server '$Server': $_"
    }
}

function Disconnect-ExchangeRemote {
    [CmdletBinding()]
    param([System.Management.Automation.Runspaces.PSSession]$Session)

    if ($Session -and $Session.State -eq 'Opened') {
        Remove-PSSession $Session
    }
}

# ─── Queue Operations ────────────────────────────────────────────────────────

function Get-ExchangeQueues {
    [CmdletBinding()]
    param(
        [string]$Server,
        [string]$Filter
    )

    $params = @{}
    if ($Server) { $params['Server'] = $Server }
    if ($Filter) { $params['Filter'] = $Filter }

    try {
        $queues = Get-Queue @params -ErrorAction Stop |
            Select-Object Identity, DeliveryType, Status, MessageCount,
                          NextHopDomain, LastError, NextRetryTime
        return $queues
    }
    catch {
        Write-Warning "Error getting queues: $_"
        return @()
    }
}

function Get-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$QueueIdentity,
        [int]$ResultSize = 100
    )

    try {
        $messages = Get-Message -Queue $QueueIdentity -ResultSize $ResultSize -ErrorAction Stop |
            Select-Object Identity, FromAddress, Status, Size, Subject,
                          DateReceived, LastError, SourceIP, SCL, RetryCount
        return $messages
    }
    catch {
        Write-Warning "Error getting messages from queue '$QueueIdentity': $_"
        return @()
    }
}

function Invoke-QueueRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$QueueIdentity
    )

    try {
        Retry-Queue -Identity $QueueIdentity -Confirm:$false -ErrorAction Stop
        return $true
    }
    catch {
        Write-Warning "Retry failed for queue '$QueueIdentity': $_"
        return $false
    }
}

function Invoke-MessageRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentity
    )

    $results = @()
    foreach ($id in $MessageIdentity) {
        try {
            Resume-Message -Identity $id -Confirm:$false -ErrorAction Stop
            $results += [PSCustomObject]@{ Identity = $id; Success = $true; Error = $null }
        }
        catch {
            $results += [PSCustomObject]@{ Identity = $id; Success = $false; Error = $_.ToString() }
        }
    }
    return $results
}

function Suspend-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentity
    )

    foreach ($id in $MessageIdentity) {
        try {
            Suspend-Message -Identity $id -Confirm:$false -ErrorAction Stop
        }
        catch {
            Write-Warning "Suspend failed for message '$id': $_"
        }
    }
}

function Remove-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentity,
        [switch]$WithNDR
    )

    foreach ($id in $MessageIdentity) {
        try {
            Remove-Message -Identity $id -WithNDR:$WithNDR -Confirm:$false -ErrorAction Stop
        }
        catch {
            Write-Warning "Remove failed for message '$id': $_"
        }
    }
}

# ─── GUI ─────────────────────────────────────────────────────────────────────

function Show-ExchangeRetryGUI {
    [CmdletBinding()]
    param()

    $script:ExSession = $null

    # ── Main Form ──
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'ExchangeRetry — Queue Manager'
    $form.Size = New-Object System.Drawing.Size(1100, 700)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(900, 500)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    # ── Top Panel (Connection) ──
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'
    $panelTop.Height = 50
    $panelTop.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 5)

    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = 'Exchange Server:'
    $lblServer.Location = New-Object System.Drawing.Point(10, 15)
    $lblServer.AutoSize = $true

    $txtServer = New-Object System.Windows.Forms.TextBox
    $txtServer.Location = New-Object System.Drawing.Point(130, 12)
    $txtServer.Size = New-Object System.Drawing.Size(250, 23)
    $txtServer.Text = $script:Config.ExchangeServer

    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text = 'Connect'
    $btnConnect.Location = New-Object System.Drawing.Point(390, 10)
    $btnConnect.Size = New-Object System.Drawing.Size(90, 28)
    $btnConnect.FlatStyle = 'Flat'

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Disconnected'
    $lblStatus.ForeColor = [System.Drawing.Color]::Gray
    $lblStatus.Location = New-Object System.Drawing.Point(490, 15)
    $lblStatus.AutoSize = $true

    $chkAutoRefresh = New-Object System.Windows.Forms.CheckBox
    $chkAutoRefresh.Text = "Auto-refresh (${($script:Config.RefreshIntervalSec)}s)"
    $chkAutoRefresh.Location = New-Object System.Drawing.Point(700, 13)
    $chkAutoRefresh.AutoSize = $true

    $panelTop.Controls.AddRange(@($lblServer, $txtServer, $btnConnect, $lblStatus, $chkAutoRefresh))

    # ── Split Container ──
    $splitContainer = New-Object System.Windows.Forms.SplitContainer
    $splitContainer.Dock = 'Fill'
    $splitContainer.Orientation = 'Horizontal'
    $splitContainer.SplitterDistance = 250
    $splitContainer.Panel1MinSize = 120
    $splitContainer.Panel2MinSize = 120

    # ── Queues Panel (Top) ──
    $lblQueues = New-Object System.Windows.Forms.Label
    $lblQueues.Text = 'Queues'
    $lblQueues.Dock = 'Top'
    $lblQueues.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $lblQueues.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 2)
    $lblQueues.AutoSize = $true

    $dgvQueues = New-Object System.Windows.Forms.DataGridView
    $dgvQueues.Dock = 'Fill'
    $dgvQueues.ReadOnly = $true
    $dgvQueues.AllowUserToAddRows = $false
    $dgvQueues.AllowUserToDeleteRows = $false
    $dgvQueues.SelectionMode = 'FullRowSelect'
    $dgvQueues.MultiSelect = $false
    $dgvQueues.AutoSizeColumnsMode = 'Fill'
    $dgvQueues.RowHeadersVisible = $false
    $dgvQueues.BackgroundColor = [System.Drawing.Color]::White
    $dgvQueues.BorderStyle = 'None'
    $dgvQueues.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 250)

    # Queue toolbar
    $panelQueueActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelQueueActions.Dock = 'Bottom'
    $panelQueueActions.Height = 38
    $panelQueueActions.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 0)

    $btnRefreshQueues = New-Object System.Windows.Forms.Button
    $btnRefreshQueues.Text = 'Refresh Queues'
    $btnRefreshQueues.Size = New-Object System.Drawing.Size(120, 28)
    $btnRefreshQueues.FlatStyle = 'Flat'

    $btnRetryQueue = New-Object System.Windows.Forms.Button
    $btnRetryQueue.Text = 'Retry Queue'
    $btnRetryQueue.Size = New-Object System.Drawing.Size(100, 28)
    $btnRetryQueue.FlatStyle = 'Flat'
    $btnRetryQueue.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRetryQueue.ForeColor = [System.Drawing.Color]::White

    $txtQueueFilter = New-Object System.Windows.Forms.TextBox
    $txtQueueFilter.Size = New-Object System.Drawing.Size(250, 23)
    $txtQueueFilter.PlaceholderText = 'Filter: Status -eq "Retry"'
    # Margin for vertical alignment in FlowLayoutPanel
    $txtQueueFilter.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 0)

    $panelQueueActions.Controls.AddRange(@($btnRefreshQueues, $btnRetryQueue, $txtQueueFilter))

    $splitContainer.Panel1.Controls.Add($dgvQueues)
    $splitContainer.Panel1.Controls.Add($panelQueueActions)
    $splitContainer.Panel1.Controls.Add($lblQueues)

    # ── Messages Panel (Bottom) ──
    $lblMessages = New-Object System.Windows.Forms.Label
    $lblMessages.Text = 'Messages in selected queue'
    $lblMessages.Dock = 'Top'
    $lblMessages.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $lblMessages.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 2)
    $lblMessages.AutoSize = $true

    $dgvMessages = New-Object System.Windows.Forms.DataGridView
    $dgvMessages.Dock = 'Fill'
    $dgvMessages.ReadOnly = $true
    $dgvMessages.AllowUserToAddRows = $false
    $dgvMessages.AllowUserToDeleteRows = $false
    $dgvMessages.SelectionMode = 'FullRowSelect'
    $dgvMessages.MultiSelect = $true
    $dgvMessages.AutoSizeColumnsMode = 'Fill'
    $dgvMessages.RowHeadersVisible = $false
    $dgvMessages.BackgroundColor = [System.Drawing.Color]::White
    $dgvMessages.BorderStyle = 'None'
    $dgvMessages.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 250, 245)

    # Message toolbar
    $panelMsgActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelMsgActions.Dock = 'Bottom'
    $panelMsgActions.Height = 38
    $panelMsgActions.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 0)

    $btnRetryMsg = New-Object System.Windows.Forms.Button
    $btnRetryMsg.Text = 'Retry Selected'
    $btnRetryMsg.Size = New-Object System.Drawing.Size(110, 28)
    $btnRetryMsg.FlatStyle = 'Flat'
    $btnRetryMsg.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $btnRetryMsg.ForeColor = [System.Drawing.Color]::White

    $btnSuspendMsg = New-Object System.Windows.Forms.Button
    $btnSuspendMsg.Text = 'Suspend Selected'
    $btnSuspendMsg.Size = New-Object System.Drawing.Size(130, 28)
    $btnSuspendMsg.FlatStyle = 'Flat'

    $btnRemoveMsg = New-Object System.Windows.Forms.Button
    $btnRemoveMsg.Text = 'Remove Selected'
    $btnRemoveMsg.Size = New-Object System.Drawing.Size(120, 28)
    $btnRemoveMsg.FlatStyle = 'Flat'
    $btnRemoveMsg.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $btnRemoveMsg.ForeColor = [System.Drawing.Color]::White

    $chkNDR = New-Object System.Windows.Forms.CheckBox
    $chkNDR.Text = 'Send NDR'
    $chkNDR.AutoSize = $true
    $chkNDR.Margin = New-Object System.Windows.Forms.Padding(10, 6, 0, 0)

    $lblMsgCount = New-Object System.Windows.Forms.Label
    $lblMsgCount.Text = ''
    $lblMsgCount.AutoSize = $true
    $lblMsgCount.Margin = New-Object System.Windows.Forms.Padding(20, 8, 0, 0)
    $lblMsgCount.ForeColor = [System.Drawing.Color]::Gray

    $panelMsgActions.Controls.AddRange(@($btnRetryMsg, $btnSuspendMsg, $btnRemoveMsg, $chkNDR, $lblMsgCount))

    $splitContainer.Panel2.Controls.Add($dgvMessages)
    $splitContainer.Panel2.Controls.Add($panelMsgActions)
    $splitContainer.Panel2.Controls.Add($lblMessages)

    # ── Status Bar ──
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready'
    $statusBar.Items.Add($statusLabel) | Out-Null

    # ── Timer for auto-refresh ──
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $script:Config.RefreshIntervalSec * 1000

    # ── Assemble Form ──
    $form.Controls.Add($splitContainer)
    $form.Controls.Add($panelTop)
    $form.Controls.Add($statusBar)

    # ─── Helper: load queues into grid ───
    $loadQueues = {
        $statusLabel.Text = 'Loading queues...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $filter = $txtQueueFilter.Text.Trim()
            $queues = Get-ExchangeQueues -Server $txtServer.Text.Trim() -Filter $filter
            $dgvQueues.DataSource = [System.Collections.ArrayList]@($queues)
            $statusLabel.Text = "Loaded $($queues.Count) queue(s)"
        }
        catch {
            $statusLabel.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to load queues: $_",
                'Error',
                'OK',
                'Error'
            )
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # ─── Helper: load messages for selected queue ───
    $loadMessages = {
        if ($dgvQueues.SelectedRows.Count -eq 0) { return }
        $queueId = $dgvQueues.SelectedRows[0].Cells['Identity'].Value.ToString()
        $lblMessages.Text = "Messages in: $queueId"
        $statusLabel.Text = 'Loading messages...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $messages = Get-QueueMessages -QueueIdentity $queueId -ResultSize $script:Config.DefaultPageSize
            $dgvMessages.DataSource = [System.Collections.ArrayList]@($messages)
            $lblMsgCount.Text = "$($messages.Count) message(s)"
            $statusLabel.Text = "Loaded $($messages.Count) message(s) from $queueId"
        }
        catch {
            $statusLabel.Text = "Error: $_"
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    # ─── Event Handlers ─────────────────────────────────────────────────
    $btnConnect.Add_Click({
        $server = $txtServer.Text.Trim()
        if (-not $server) {
            [System.Windows.Forms.MessageBox]::Show('Enter Exchange server name.', 'Warning', 'OK', 'Warning')
            return
        }
        $statusLabel.Text = "Connecting to $server..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if ($script:ExSession) {
                Disconnect-ExchangeRemote -Session $script:ExSession
            }
            $script:ExSession = Connect-ExchangeRemote -Server $server
            $lblStatus.Text = "Connected to $server"
            $lblStatus.ForeColor = [System.Drawing.Color]::Green
            $btnConnect.Text = 'Reconnect'
            $statusLabel.Text = "Connected to $server"
            & $loadQueues
        }
        catch {
            $lblStatus.Text = 'Connection failed'
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $statusLabel.Text = "Connection failed: $_"
            [System.Windows.Forms.MessageBox]::Show(
                "Connection failed:`n$_",
                'Connection Error',
                'OK',
                'Error'
            )
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

    $btnRefreshQueues.Add_Click({ & $loadQueues })

    $dgvQueues.Add_SelectionChanged({ & $loadMessages })

    $btnRetryQueue.Add_Click({
        if ($dgvQueues.SelectedRows.Count -eq 0) { return }
        $queueId = $dgvQueues.SelectedRows[0].Cells['Identity'].Value.ToString()
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Retry all messages in queue '$queueId'?",
            'Confirm Retry',
            'YesNo',
            'Question'
        )
        if ($confirm -eq 'Yes') {
            $result = Invoke-QueueRetry -QueueIdentity $queueId
            if ($result) {
                $statusLabel.Text = "Queue '$queueId' retry initiated"
            }
            & $loadQueues
        }
    })

    $btnRetryMsg.Add_Click({
        $selected = @()
        foreach ($row in $dgvMessages.SelectedRows) {
            $selected += $row.Cells['Identity'].Value.ToString()
        }
        if ($selected.Count -eq 0) { return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Retry $($selected.Count) message(s)?",
            'Confirm Retry',
            'YesNo',
            'Question'
        )
        if ($confirm -eq 'Yes') {
            $results = Invoke-MessageRetry -MessageIdentity $selected
            $ok = ($results | Where-Object Success).Count
            $fail = ($results | Where-Object { -not $_.Success }).Count
            $statusLabel.Text = "Retry: $ok succeeded, $fail failed"
            & $loadMessages
        }
    })

    $btnSuspendMsg.Add_Click({
        $selected = @()
        foreach ($row in $dgvMessages.SelectedRows) {
            $selected += $row.Cells['Identity'].Value.ToString()
        }
        if ($selected.Count -eq 0) { return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Suspend $($selected.Count) message(s)?",
            'Confirm Suspend',
            'YesNo',
            'Question'
        )
        if ($confirm -eq 'Yes') {
            Suspend-QueueMessages -MessageIdentity $selected
            $statusLabel.Text = "Suspended $($selected.Count) message(s)"
            & $loadMessages
        }
    })

    $btnRemoveMsg.Add_Click({
        $selected = @()
        foreach ($row in $dgvMessages.SelectedRows) {
            $selected += $row.Cells['Identity'].Value.ToString()
        }
        if ($selected.Count -eq 0) { return }
        $withNDR = $chkNDR.Checked
        $ndrText = if ($withNDR) { ' with NDR' } else { ' without NDR' }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "PERMANENTLY remove $($selected.Count) message(s)$ndrText?`n`nThis action cannot be undone!",
            'Confirm Remove',
            'YesNo',
            'Warning'
        )
        if ($confirm -eq 'Yes') {
            Remove-QueueMessages -MessageIdentity $selected -WithNDR:$withNDR
            $statusLabel.Text = "Removed $($selected.Count) message(s)"
            & $loadMessages
        }
    })

    $chkAutoRefresh.Add_CheckedChanged({
        if ($chkAutoRefresh.Checked) {
            $timer.Start()
        } else {
            $timer.Stop()
        }
    })

    $timer.Add_Tick({ & $loadQueues })

    $form.Add_FormClosing({
        $timer.Stop()
        $timer.Dispose()
        if ($script:ExSession) {
            Disconnect-ExchangeRemote -Session $script:ExSession
        }
    })

    # ── Show ──
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# ─── Entry Point ─────────────────────────────────────────────────────────────

if ($MyInvocation.InvocationName -ne '.') {
    Show-ExchangeRetryGUI
}
