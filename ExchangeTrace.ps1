<#
.SYNOPSIS
    ExchangeTrace — консольная утилита для трассировки писем в Exchange.
.DESCRIPTION
    Парсит заголовки писем (Received, X-MS-Exchange-*), ищет по транспортным логам,
    строит маршрут прохождения письма через серверы, генерирует отчёты.
.NOTES
    Version: 0.1.0
.EXAMPLE
    .\ExchangeTrace.ps1 -HeaderFile .\header.txt
    .\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01
    .\ExchangeTrace.ps1 -Sender user@domain.com -StartDate "2026-03-01" -Server exchange01
    .\ExchangeTrace.ps1 -TransportLogPath "\\exchange01\TransportLogs" -SearchPattern "user@domain.com"
    .\ExchangeTrace.ps1 -Report Full -Server exchange01
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'Header')]
param(
    [Parameter(ParameterSetName = 'Header')]
    [string]$HeaderFile,

    [Parameter(ParameterSetName = 'Header')]
    [string]$HeaderText,

    [Parameter(ParameterSetName = 'Track')]
    [string]$MessageId,

    [Parameter(ParameterSetName = 'Track')]
    [Parameter(ParameterSetName = 'Search')]
    [string]$Sender,

    [Parameter(ParameterSetName = 'Track')]
    [Parameter(ParameterSetName = 'Search')]
    [string]$Recipient,

    [Parameter(ParameterSetName = 'Track')]
    [Parameter(ParameterSetName = 'Search')]
    [Parameter(ParameterSetName = 'Report')]
    [string]$Server,

    [Parameter(ParameterSetName = 'Track')]
    [Parameter(ParameterSetName = 'Search')]
    [datetime]$StartDate = (Get-Date).AddDays(-1),

    [Parameter(ParameterSetName = 'Track')]
    [Parameter(ParameterSetName = 'Search')]
    [datetime]$EndDate = (Get-Date),

    [Parameter(ParameterSetName = 'LogSearch')]
    [string]$TransportLogPath,

    [Parameter(ParameterSetName = 'LogSearch')]
    [string]$SearchPattern,

    [Parameter(ParameterSetName = 'Report')]
    [ValidateSet('Full', 'Queues', 'Connectors', 'AgentLog', 'RoutingTable', 'DSN', 'Summary')]
    [string]$Report,

    [string]$OutputFile,

    [ValidateSet('Table', 'List', 'CSV', 'JSON')]
    [string]$OutputFormat = 'Table'
)

# ─── Header Parsing ──────────────────────────────────────────────────────────

function Parse-EmailHeaders {
    <#
    .SYNOPSIS
        Парсит заголовки email и извлекает маршрут прохождения.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RawHeaders
    )

    $result = [PSCustomObject]@{
        Hops           = @()
        MessageId      = $null
        From           = $null
        To             = $null
        Subject        = $null
        Date           = $null
        ContentType    = $null
        SPF            = $null
        DKIM           = $null
        DMARC          = $null
        XHeaders       = @{}
        TotalHops      = 0
        TotalDelayMs   = 0
    }

    # Extract simple headers
    if ($RawHeaders -match '(?m)^Message-ID:\s*(.+)$') {
        $result.MessageId = $Matches[1].Trim()
    }
    if ($RawHeaders -match '(?m)^From:\s*(.+)$') {
        $result.From = $Matches[1].Trim()
    }
    if ($RawHeaders -match '(?m)^To:\s*(.+)$') {
        $result.To = $Matches[1].Trim()
    }
    if ($RawHeaders -match '(?m)^Subject:\s*(.+)$') {
        $result.Subject = $Matches[1].Trim()
    }
    if ($RawHeaders -match '(?m)^Date:\s*(.+)$') {
        $result.Date = $Matches[1].Trim()
    }

    # Extract authentication results
    if ($RawHeaders -match '(?m)spf=(\w+)') {
        $result.SPF = $Matches[1]
    }
    if ($RawHeaders -match '(?m)dkim=(\w+)') {
        $result.DKIM = $Matches[1]
    }
    if ($RawHeaders -match '(?m)dmarc=(\w+)') {
        $result.DMARC = $Matches[1]
    }

    # Extract X-MS-Exchange-* and other X- headers
    $xHeaderMatches = [regex]::Matches($RawHeaders, '(?m)^(X-[\w-]+):\s*(.+)$')
    foreach ($match in $xHeaderMatches) {
        $headerName = $match.Groups[1].Value
        $headerValue = $match.Groups[2].Value.Trim()
        $result.XHeaders[$headerName] = $headerValue
    }

    # Parse Received headers (bottom-to-top = chronological order)
    # Unfold multi-line Received headers
    $unfolded = $RawHeaders -replace '(\r?\n)\s+', ' '
    $receivedMatches = [regex]::Matches($unfolded, '(?m)^Received:\s*(.+?)(?=^[\w-]+:|\z)', [System.Text.RegularExpressions.RegexOptions]::Multiline)

    $hops = @()
    foreach ($match in $receivedMatches) {
        $receivedLine = $match.Groups[1].Value.Trim()

        $hop = [PSCustomObject]@{
            Raw       = $receivedLine
            From      = $null
            By        = $null
            With      = $null
            Timestamp = $null
            Delay     = $null
        }

        # Parse "from <server>"
        if ($receivedLine -match 'from\s+([\w\.\-]+(?:\s*\([^\)]*\))?)') {
            $hop.From = $Matches[1].Trim()
        }

        # Parse "by <server>"
        if ($receivedLine -match 'by\s+([\w\.\-]+(?:\s*\([^\)]*\))?)') {
            $hop.By = $Matches[1].Trim()
        }

        # Parse "with <protocol>"
        if ($receivedLine -match 'with\s+(\w+)') {
            $hop.With = $Matches[1]
        }

        # Parse timestamp at the end (after ;)
        if ($receivedLine -match ';\s*(.+)$') {
            $tsString = $Matches[1].Trim()
            try {
                $hop.Timestamp = [datetime]::Parse($tsString)
            }
            catch {
                $hop.Timestamp = $tsString
            }
        }

        $hops += $hop
    }

    # Reverse to chronological order (earliest first)
    [Array]::Reverse($hops)

    # Calculate delays between hops
    for ($i = 1; $i -lt $hops.Count; $i++) {
        if ($hops[$i].Timestamp -is [datetime] -and $hops[$i - 1].Timestamp -is [datetime]) {
            $delay = ($hops[$i].Timestamp - $hops[$i - 1].Timestamp)
            $hops[$i].Delay = $delay
        }
    }

    $result.Hops = $hops
    $result.TotalHops = $hops.Count
    if ($hops.Count -ge 2 -and $hops[0].Timestamp -is [datetime] -and $hops[-1].Timestamp -is [datetime]) {
        $result.TotalDelayMs = ($hops[-1].Timestamp - $hops[0].Timestamp).TotalMilliseconds
    }

    return $result
}

function Format-HeaderTrace {
    <#
    .SYNOPSIS
        Красиво выводит маршрут письма.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $ParsedHeaders
    )

    Write-Host "`n═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  EMAIL TRACE REPORT" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    Write-Host "`n  Message-ID : $($ParsedHeaders.MessageId)" -ForegroundColor White
    Write-Host "  From       : $($ParsedHeaders.From)"
    Write-Host "  To         : $($ParsedHeaders.To)"
    Write-Host "  Subject    : $($ParsedHeaders.Subject)"
    Write-Host "  Date       : $($ParsedHeaders.Date)"

    # Authentication
    Write-Host "`n── Authentication ─────────────────────────────────────────────" -ForegroundColor Yellow
    $spfColor = if ($ParsedHeaders.SPF -eq 'pass') { 'Green' } else { 'Red' }
    $dkimColor = if ($ParsedHeaders.DKIM -eq 'pass') { 'Green' } else { 'Red' }
    $dmarcColor = if ($ParsedHeaders.DMARC -eq 'pass') { 'Green' } else { 'Red' }
    Write-Host "  SPF   : $(if ($ParsedHeaders.SPF) { $ParsedHeaders.SPF } else { 'N/A' })" -ForegroundColor $spfColor
    Write-Host "  DKIM  : $(if ($ParsedHeaders.DKIM) { $ParsedHeaders.DKIM } else { 'N/A' })" -ForegroundColor $dkimColor
    Write-Host "  DMARC : $(if ($ParsedHeaders.DMARC) { $ParsedHeaders.DMARC } else { 'N/A' })" -ForegroundColor $dmarcColor

    # Route
    Write-Host "`n── Route ($($ParsedHeaders.TotalHops) hops, total: $([math]::Round($ParsedHeaders.TotalDelayMs / 1000, 2))s) ──" -ForegroundColor Yellow
    $hopNum = 0
    foreach ($hop in $ParsedHeaders.Hops) {
        $hopNum++
        $delayStr = ''
        if ($hop.Delay) {
            $delaySec = [math]::Round($hop.Delay.TotalSeconds, 2)
            $delayColor = if ($delaySec -gt 5) { 'Red' } elseif ($delaySec -gt 1) { 'Yellow' } else { 'Green' }
            $delayStr = " [+${delaySec}s]"
        }

        Write-Host "  [$hopNum] " -NoNewline -ForegroundColor DarkGray
        Write-Host "$(if ($hop.From) { $hop.From } else { '?' })" -NoNewline -ForegroundColor White
        Write-Host " → " -NoNewline -ForegroundColor DarkGray
        Write-Host "$(if ($hop.By) { $hop.By } else { '?' })" -NoNewline -ForegroundColor Green
        Write-Host " ($(if ($hop.With) { $hop.With } else { '?' }))" -NoNewline -ForegroundColor DarkGray
        if ($delayStr) {
            Write-Host $delayStr -ForegroundColor $delayColor
        } else {
            Write-Host ''
        }
        if ($hop.Timestamp -is [datetime]) {
            Write-Host "       $($hop.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor DarkGray
        }
    }

    # X-Headers
    if ($ParsedHeaders.XHeaders.Count -gt 0) {
        Write-Host "`n── Exchange X-Headers ─────────────────────────────────────────" -ForegroundColor Yellow
        foreach ($key in ($ParsedHeaders.XHeaders.Keys | Sort-Object)) {
            Write-Host "  $key : $($ParsedHeaders.XHeaders[$key])" -ForegroundColor Gray
        }
    }

    Write-Host "`n═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan
}

# ─── Transport Log Search ────────────────────────────────────────────────────

function Search-TransportLogs {
    <#
    .SYNOPSIS
        Поиск по текстовым транспортным логам Exchange (SMTP Send/Receive, Message Tracking).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [int]$ContextLines = 2
    )

    if (-not (Test-Path $LogPath)) {
        Write-Error "Log path not found: $LogPath"
        return
    }

    $logFiles = Get-ChildItem -Path $LogPath -Filter '*.log' -Recurse |
        Sort-Object LastWriteTime -Descending

    if ($logFiles.Count -eq 0) {
        Write-Warning "No .log files found in $LogPath"
        return
    }

    Write-Host "Searching $($logFiles.Count) log files for '$Pattern'..." -ForegroundColor Cyan

    $allMatches = @()

    foreach ($file in $logFiles) {
        $lines = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match [regex]::Escape($Pattern)) {
                $startLine = [Math]::Max(0, $i - $ContextLines)
                $endLine = [Math]::Min($lines.Count - 1, $i + $ContextLines)
                $context = $lines[$startLine..$endLine] -join "`n"

                $allMatches += [PSCustomObject]@{
                    File      = $file.Name
                    FilePath  = $file.FullName
                    Line      = $i + 1
                    Match     = $lines[$i]
                    Context   = $context
                    Timestamp = $file.LastWriteTime
                }
            }
        }
    }

    Write-Host "Found $($allMatches.Count) match(es) across $($logFiles.Count) file(s)" -ForegroundColor Green
    return $allMatches
}

# ─── Message Tracking (Exchange cmdlets) ─────────────────────────────────────

function Trace-ExchangeMessage {
    <#
    .SYNOPSIS
        Трассировка сообщения через Get-MessageTrackingLog.
    #>
    [CmdletBinding()]
    param(
        [string]$MessageId,
        [string]$Sender,
        [string]$Recipient,
        [string]$Server,
        [datetime]$Start = (Get-Date).AddDays(-1),
        [datetime]$End = (Get-Date)
    )

    $params = @{
        Start       = $Start
        End         = $End
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    }

    if ($MessageId)  { $params['MessageId'] = $MessageId }
    if ($Sender)     { $params['Sender'] = $Sender }
    if ($Recipient)  { $params['Recipients'] = $Recipient }
    if ($Server)     { $params['Server'] = $Server }

    try {
        $logs = Get-MessageTrackingLog @params |
            Select-Object Timestamp, EventId, Source, Sender,
                          @{N='Recipients';E={$_.Recipients -join '; '}},
                          MessageSubject, ServerHostname, ServerIp,
                          ConnectorId, SourceContext, MessageId,
                          TotalBytes, RecipientCount, RecipientStatus,
                          @{N='InternalMessageId';E={$_.InternalMessageId}} |
            Sort-Object Timestamp

        return $logs
    }
    catch {
        Write-Error "Message tracking failed: $_"
        return @()
    }
}

function Format-MessageTrace {
    <#
    .SYNOPSIS
        Красивый вывод трассировки сообщения.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $TrackingResults
    )

    if ($TrackingResults.Count -eq 0) {
        Write-Host "No tracking results found." -ForegroundColor Yellow
        return
    }

    Write-Host "`n═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  MESSAGE TRACKING TRACE ($($TrackingResults.Count) events)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    $first = $TrackingResults[0]
    Write-Host "  Subject    : $($first.MessageSubject)"
    Write-Host "  Message-ID : $($first.MessageId)"
    Write-Host "  Sender     : $($first.Sender)"
    Write-Host ""

    foreach ($event in $TrackingResults) {
        $eventColor = switch ($event.EventId) {
            'RECEIVE'   { 'Green' }
            'SEND'      { 'Cyan' }
            'DELIVER'   { 'Green' }
            'FAIL'      { 'Red' }
            'DSN'       { 'Red' }
            'DEFER'     { 'Yellow' }
            'EXPAND'    { 'DarkCyan' }
            'REDIRECT'  { 'Magenta' }
            'RESOLVE'   { 'DarkGray' }
            'SUBMIT'    { 'White' }
            default     { 'Gray' }
        }

        $ts = $event.Timestamp.ToString('HH:mm:ss.fff')
        Write-Host "  $ts " -NoNewline -ForegroundColor DarkGray
        Write-Host "$($event.EventId.PadRight(12))" -NoNewline -ForegroundColor $eventColor
        Write-Host "$($event.Source.PadRight(10))" -NoNewline -ForegroundColor Gray
        Write-Host "$($event.ServerHostname)" -NoNewline -ForegroundColor White
        if ($event.ConnectorId) {
            Write-Host " [$($event.ConnectorId)]" -NoNewline -ForegroundColor DarkGray
        }
        Write-Host ''
        if ($event.Recipients) {
            Write-Host "             → $($event.Recipients)" -ForegroundColor DarkCyan
        }
        if ($event.RecipientStatus -and $event.EventId -in @('FAIL', 'DSN', 'DEFER')) {
            Write-Host "             ! $($event.RecipientStatus)" -ForegroundColor Red
        }
    }

    Write-Host "`n═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan
}

# ─── Transport Reports ───────────────────────────────────────────────────────

function Get-TransportReport {
    <#
    .SYNOPSIS
        Генерация отчётов по транспортной подсистеме Exchange.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Full', 'Queues', 'Connectors', 'AgentLog', 'RoutingTable', 'DSN', 'Summary')]
        [string]$ReportType,

        [string]$Server
    )

    $serverParam = @{}
    if ($Server) { $serverParam['Server'] = $Server }

    $report = [ordered]@{}

    # ── Queue Status ──
    if ($ReportType -in @('Full', 'Queues', 'Summary')) {
        Write-Host "`n── Queue Status ───────────────────────────────────────────────" -ForegroundColor Yellow
        try {
            $queues = Get-Queue @serverParam -ErrorAction Stop
            $report['Queues'] = $queues

            $queueSummary = $queues | Group-Object Status | Select-Object Name, Count
            Write-Host "  Total queues: $($queues.Count)"
            foreach ($g in $queueSummary) {
                $color = if ($g.Name -eq 'Ready') { 'Green' } elseif ($g.Name -eq 'Retry') { 'Red' } else { 'Yellow' }
                Write-Host "    $($g.Name): $($g.Count)" -ForegroundColor $color
            }

            $totalMessages = ($queues | Measure-Object -Property MessageCount -Sum).Sum
            Write-Host "  Total messages in queues: $totalMessages"

            $retryQueues = $queues | Where-Object { $_.Status -eq 'Retry' }
            if ($retryQueues) {
                Write-Host "`n  Retry queues:" -ForegroundColor Red
                foreach ($q in $retryQueues) {
                    Write-Host "    $($q.Identity) — $($q.MessageCount) msgs — $($q.LastError)" -ForegroundColor Red
                }
            }
        }
        catch {
            Write-Warning "  Cannot get queues: $_"
        }
    }

    # ── Send/Receive Connectors ──
    if ($ReportType -in @('Full', 'Connectors')) {
        Write-Host "`n── Send Connectors ────────────────────────────────────────────" -ForegroundColor Yellow
        try {
            $sendConn = Get-SendConnector -ErrorAction Stop
            $report['SendConnectors'] = $sendConn
            foreach ($c in $sendConn) {
                $state = if ($c.Enabled) { '[ON]' } else { '[OFF]' }
                $stateColor = if ($c.Enabled) { 'Green' } else { 'Red' }
                Write-Host "  $state " -NoNewline -ForegroundColor $stateColor
                Write-Host "$($c.Name) → $($c.AddressSpaces -join ', ')" -ForegroundColor White
                Write-Host "         SmartHosts: $($c.SmartHosts -join ', ')" -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Warning "  Cannot get send connectors: $_"
        }

        Write-Host "`n── Receive Connectors ─────────────────────────────────────────" -ForegroundColor Yellow
        try {
            $recvConn = Get-ReceiveConnector @serverParam -ErrorAction Stop
            $report['ReceiveConnectors'] = $recvConn
            foreach ($c in $recvConn) {
                $state = if ($c.Enabled) { '[ON]' } else { '[OFF]' }
                $stateColor = if ($c.Enabled) { 'Green' } else { 'Red' }
                Write-Host "  $state " -NoNewline -ForegroundColor $stateColor
                Write-Host "$($c.Name) — $($c.Bindings -join ', ')" -ForegroundColor White
                Write-Host "         Auth: $($c.AuthMechanism)" -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Warning "  Cannot get receive connectors: $_"
        }
    }

    # ── Transport Agent Log ──
    if ($ReportType -in @('Full', 'AgentLog')) {
        Write-Host "`n── Transport Agent Log (last 50) ──────────────────────────────" -ForegroundColor Yellow
        try {
            $agentLog = Get-AgentLog @serverParam -ErrorAction Stop |
                Select-Object -Last 50 Timestamp, SessionId, IPAddress, MessageId, P1FromAddress, P2FromAddresses, Agent, Event, Action, SmtpResponse
            $report['AgentLog'] = $agentLog

            foreach ($entry in $agentLog) {
                $actionColor = if ($entry.Action -eq 'RejectMessage') { 'Red' } elseif ($entry.Action -eq 'AcceptMessage') { 'Green' } else { 'Yellow' }
                Write-Host "  $($entry.Timestamp.ToString('HH:mm:ss')) $($entry.Agent.PadRight(30)) " -NoNewline -ForegroundColor DarkGray
                Write-Host "$($entry.Action)" -ForegroundColor $actionColor
            }
        }
        catch {
            Write-Warning "  Cannot get agent log: $_"
        }
    }

    # ── Routing Table ──
    if ($ReportType -in @('Full', 'RoutingTable')) {
        Write-Host "`n── Routing Table ──────────────────────────────────────────────" -ForegroundColor Yellow
        try {
            $config = Get-TransportConfig -ErrorAction Stop
            $report['TransportConfig'] = $config
            Write-Host "  Max receive size      : $($config.MaxReceiveSize)"
            Write-Host "  Max send size         : $($config.MaxSendSize)"
            Write-Host "  Max recipients        : $($config.MaxRecipientEnvelopeLimit)"
            Write-Host "  Shadow redundancy     : $($config.ShadowRedundancyEnabled)"
            Write-Host "  Safety net hold time  : $($config.SafetyNetHoldTime)"

            $transportServer = Get-TransportService @serverParam -ErrorAction Stop
            $report['TransportService'] = $transportServer
            Write-Host "`n  Transport Service:"
            Write-Host "    Message tracking   : $($transportServer.MessageTrackingLogEnabled)"
            Write-Host "    Log path           : $($transportServer.MessageTrackingLogPath)"
            Write-Host "    Connectivity log   : $($transportServer.ConnectivityLogPath)"
            Write-Host "    Send protocol log  : $($transportServer.SendProtocolLogPath)"
            Write-Host "    Recv protocol log  : $($transportServer.ReceiveProtocolLogPath)"
        }
        catch {
            Write-Warning "  Cannot get transport config: $_"
        }
    }

    # ── DSN (Delivery Status Notifications) ──
    if ($ReportType -in @('Full', 'DSN')) {
        Write-Host "`n── DSN Configuration ──────────────────────────────────────────" -ForegroundColor Yellow
        try {
            $config = Get-TransportConfig -ErrorAction Stop
            Write-Host "  Generate copy of DSN to : $($config.GenerateCopyOfDSNFor -join ', ')"
            Write-Host "  External DSN language   : $($config.ExternalDsnLanguageDetectionEnabled)"
            Write-Host "  Internal DSN language   : $($config.InternalDsnLanguageDetectionEnabled)"
            Write-Host "  External postmaster     : $($config.ExternalPostmasterAddress)"

            # Recent DSN/NDR from tracking
            Write-Host "`n  Recent DSN events (last 24h):" -ForegroundColor Yellow
            $dsnParams = @{
                EventId    = 'DSN'
                Start      = (Get-Date).AddDays(-1)
                End        = Get-Date
                ResultSize = 50
            }
            if ($Server) { $dsnParams['Server'] = $Server }
            $dsnEvents = Get-MessageTrackingLog @dsnParams -ErrorAction Stop
            $report['DSNEvents'] = $dsnEvents

            if ($dsnEvents.Count -eq 0) {
                Write-Host "    No DSN events in the last 24h" -ForegroundColor Green
            }
            else {
                foreach ($dsn in $dsnEvents) {
                    Write-Host "    $($dsn.Timestamp.ToString('MM-dd HH:mm')) $($dsn.Sender) → $($dsn.Recipients -join ', ')" -ForegroundColor Red
                    Write-Host "      $($dsn.RecipientStatus)" -ForegroundColor DarkRed
                }
            }
        }
        catch {
            Write-Warning "  Cannot get DSN config: $_"
        }
    }

    # ── Summary ──
    if ($ReportType -eq 'Summary') {
        Write-Host "`n── Delivery Summary (last 24h) ────────────────────────────────" -ForegroundColor Yellow
        try {
            $trackParams = @{
                Start      = (Get-Date).AddDays(-1)
                End        = Get-Date
                ResultSize = 'Unlimited'
            }
            if ($Server) { $trackParams['Server'] = $Server }
            $allEvents = Get-MessageTrackingLog @trackParams -ErrorAction Stop

            $summary = $allEvents | Group-Object EventId | Select-Object Name, Count | Sort-Object Count -Descending
            $report['Summary'] = $summary

            foreach ($s in $summary) {
                $color = switch ($s.Name) {
                    'DELIVER' { 'Green' }
                    'FAIL'    { 'Red' }
                    'DEFER'   { 'Yellow' }
                    'DSN'     { 'Red' }
                    default   { 'White' }
                }
                Write-Host "    $($s.Name.PadRight(15)) $($s.Count)" -ForegroundColor $color
            }

            $delivered = ($allEvents | Where-Object EventId -eq 'DELIVER').Count
            $failed = ($allEvents | Where-Object EventId -eq 'FAIL').Count
            $total = $delivered + $failed
            if ($total -gt 0) {
                $rate = [math]::Round(($delivered / $total) * 100, 1)
                Write-Host "`n    Delivery rate: $rate% ($delivered/$total)" -ForegroundColor $(if ($rate -ge 95) { 'Green' } elseif ($rate -ge 80) { 'Yellow' } else { 'Red' })
            }
        }
        catch {
            Write-Warning "  Cannot get tracking summary: $_"
        }
    }

    Write-Host ""
    return $report
}

# ─── Output Helpers ──────────────────────────────────────────────────────────

function Export-Results {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        [string]$FilePath,
        [string]$Format = 'CSV'
    )

    switch ($Format) {
        'CSV'  { $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 }
        'JSON' { $Data | ConvertTo-Json -Depth 5 | Set-Content -Path $FilePath -Encoding UTF8 }
        default { $Data | Out-File -FilePath $FilePath -Encoding UTF8 }
    }
    Write-Host "Results exported to: $FilePath" -ForegroundColor Green
}

# ─── Main Entry Point ────────────────────────────────────────────────────────

switch ($PSCmdlet.ParameterSetName) {

    'Header' {
        # Parse email headers
        $rawHeaders = $null
        if ($HeaderFile) {
            if (-not (Test-Path $HeaderFile)) {
                Write-Error "Header file not found: $HeaderFile"
                exit 1
            }
            $rawHeaders = Get-Content -Path $HeaderFile -Raw
        }
        elseif ($HeaderText) {
            $rawHeaders = $HeaderText
        }
        else {
            Write-Host "Paste email headers below (end with empty line):" -ForegroundColor Cyan
            $lines = @()
            while ($true) {
                $line = Read-Host
                if ([string]::IsNullOrEmpty($line)) { break }
                $lines += $line
            }
            $rawHeaders = $lines -join "`n"
        }

        if ($rawHeaders) {
            $parsed = Parse-EmailHeaders -RawHeaders $rawHeaders
            Format-HeaderTrace -ParsedHeaders $parsed

            if ($OutputFile) {
                Export-Results -Data $parsed.Hops -FilePath $OutputFile -Format $OutputFormat
            }
        }
    }

    'Track' {
        # Message tracking via Exchange cmdlets
        $traceParams = @{
            Start = $StartDate
            End   = $EndDate
        }
        if ($MessageId) { $traceParams['MessageId'] = $MessageId }
        if ($Sender)    { $traceParams['Sender'] = $Sender }
        if ($Recipient) { $traceParams['Recipient'] = $Recipient }
        if ($Server)    { $traceParams['Server'] = $Server }

        $results = Trace-ExchangeMessage @traceParams
        Format-MessageTrace -TrackingResults $results

        if ($OutputFile) {
            Export-Results -Data $results -FilePath $OutputFile -Format $OutputFormat
        }
    }

    'Search' {
        # Search by sender/recipient
        $traceParams = @{
            Start = $StartDate
            End   = $EndDate
        }
        if ($Sender)    { $traceParams['Sender'] = $Sender }
        if ($Recipient) { $traceParams['Recipient'] = $Recipient }
        if ($Server)    { $traceParams['Server'] = $Server }

        $results = Trace-ExchangeMessage @traceParams
        Format-MessageTrace -TrackingResults $results

        if ($OutputFile) {
            Export-Results -Data $results -FilePath $OutputFile -Format $OutputFormat
        }
    }

    'LogSearch' {
        # Raw transport log search
        if (-not $TransportLogPath -or -not $SearchPattern) {
            Write-Error "Both -TransportLogPath and -SearchPattern are required."
            exit 1
        }

        $matches = Search-TransportLogs -LogPath $TransportLogPath -Pattern $SearchPattern
        if ($matches) {
            foreach ($m in $matches) {
                Write-Host "`n── $($m.File):$($m.Line) ──" -ForegroundColor Yellow
                Write-Host $m.Context
            }

            if ($OutputFile) {
                Export-Results -Data $matches -FilePath $OutputFile -Format $OutputFormat
            }
        }
    }

    'Report' {
        # Transport reports
        $reportData = Get-TransportReport -ReportType $Report -Server $Server

        if ($OutputFile -and $reportData) {
            $reportData | ConvertTo-Json -Depth 5 | Set-Content -Path $OutputFile -Encoding UTF8
            Write-Host "Report exported to: $OutputFile" -ForegroundColor Green
        }
    }
}
