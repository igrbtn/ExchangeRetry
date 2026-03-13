<#
.SYNOPSIS
    ExchangeRetry — полнофункциональный GUI для мониторинга транспорта Microsoft Exchange.
.DESCRIPTION
    Отслеживание сообщений на каждом этапе транспортного пайплайна:
    1. Dashboard — здоровье транспорта, очереди, доставка, ошибки
    2. Queues — управление очередями (retry/suspend/remove)
    3. Message Tracking — трассировка с фильтрами по EventId/Server/Connector
    4. Protocol Logs — парсинг SMTP Send/Receive protocol logs
    5. Transport Logs — Connectivity, Agent, Routing logs
    6. Header Analyzer — парсинг заголовков email
    7. Reports — все отчёты по транспортной подсистеме
.NOTES
    Version: 0.3.0
    Requires: Exchange Management Shell (EMS) или удалённое подключение к Exchange.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ─── Configuration ───────────────────────────────────────────────────────────

$script:Config = @{
    ExchangeServer     = $env:EXCHANGE_SERVER
    DefaultPageSize    = 200
    RefreshIntervalSec = 30
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — EXCHANGE CONNECTION
# ═══════════════════════════════════════════════════════════════════════════════

function Connect-ExchangeRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Server,
        [PSCredential]$Credential
    )
    $uri = "http://$Server/PowerShell/"
    $p = @{ ConfigurationName = 'Microsoft.Exchange'; ConnectionUri = $uri; Authentication = 'Kerberos' }
    if ($Credential) { $p['Credential'] = $Credential }
    try {
        $session = New-PSSession @p -ErrorAction Stop
        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        return $session
    }
    catch { throw "Failed to connect to '$Server': $_" }
}

function Disconnect-ExchangeRemote {
    [CmdletBinding()]
    param([System.Management.Automation.Runspaces.PSSession]$Session)
    if ($Session -and $Session.State -eq 'Opened') { Remove-PSSession $Session }
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — QUEUE OPERATIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Get-ExchangeQueues {
    [CmdletBinding()]
    param([string]$Server, [string]$Filter)
    $p = @{}
    if ($Server) { $p['Server'] = $Server }
    if ($Filter) { $p['Filter'] = $Filter }
    try {
        return Get-Queue @p -ErrorAction Stop |
            Select-Object Identity, DeliveryType, Status, MessageCount,
                          NextHopDomain, LastError, NextRetryTime
    }
    catch { Write-Warning "Error getting queues: $_"; return @() }
}

function Get-QueueMessages {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueIdentity, [int]$ResultSize = 200)
    try {
        return Get-Message -Queue $QueueIdentity -ResultSize $ResultSize -ErrorAction Stop |
            Select-Object Identity, FromAddress, Status, Size, Subject,
                          DateReceived, LastError, SourceIP, SCL, RetryCount
    }
    catch { Write-Warning "Error getting messages: $_"; return @() }
}

function Invoke-QueueRetry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$QueueIdentity)
    try { Retry-Queue -Identity $QueueIdentity -Confirm:$false -ErrorAction Stop; return $true }
    catch { Write-Warning "Retry failed: $_"; return $false }
}

function Invoke-MessageRetry {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$MessageIdentity)
    $results = @()
    foreach ($id in $MessageIdentity) {
        try {
            Resume-Message -Identity $id -Confirm:$false -ErrorAction Stop
            $results += [PSCustomObject]@{ Identity = $id; Success = $true; Error = $null }
        }
        catch { $results += [PSCustomObject]@{ Identity = $id; Success = $false; Error = $_.ToString() } }
    }
    return $results
}

function Suspend-QueueMessages {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$MessageIdentity)
    foreach ($id in $MessageIdentity) {
        try { Suspend-Message -Identity $id -Confirm:$false -ErrorAction Stop }
        catch { Write-Warning "Suspend failed for '$id': $_" }
    }
}

function Remove-QueueMessages {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$MessageIdentity, [switch]$WithNDR)
    foreach ($id in $MessageIdentity) {
        try { Remove-Message -Identity $id -WithNDR:$WithNDR -Confirm:$false -ErrorAction Stop }
        catch { Write-Warning "Remove failed for '$id': $_" }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — MESSAGE TRACKING
# ═══════════════════════════════════════════════════════════════════════════════

function Trace-ExchangeMessage {
    [CmdletBinding()]
    param(
        [string]$MessageId, [string]$Sender, [string]$Recipient,
        [string]$Server, [string]$EventId, [string]$ConnectorId,
        [string]$Source, [string]$Subject,
        [datetime]$Start = (Get-Date).AddDays(-1),
        [datetime]$End = (Get-Date),
        [int]$ResultSize = 1000
    )
    $p = @{ Start = $Start; End = $End; ErrorAction = 'Stop' }
    if ($ResultSize -gt 0) { $p['ResultSize'] = $ResultSize } else { $p['ResultSize'] = 'Unlimited' }
    if ($MessageId)   { $p['MessageId'] = $MessageId }
    if ($Sender)      { $p['Sender'] = $Sender }
    if ($Recipient)   { $p['Recipients'] = $Recipient }
    if ($Server)      { $p['Server'] = $Server }
    if ($EventId)     { $p['EventId'] = $EventId }
    if ($ConnectorId) { $p['ConnectorId'] = $ConnectorId }
    if ($Source)      { $p['Source'] = $Source }

    try {
        $logs = Get-MessageTrackingLog @p |
            Select-Object Timestamp, EventId, Source, Sender,
                @{N='Recipients';E={$_.Recipients -join '; '}},
                MessageSubject, ServerHostname, ServerIp,
                ConnectorId, SourceContext, MessageId,
                TotalBytes, RecipientCount, RecipientStatus,
                @{N='InternalMessageId';E={$_.InternalMessageId}},
                ClientIp, ClientHostname, OriginalClientIp |
            Sort-Object Timestamp

        if ($Subject) {
            $logs = $logs | Where-Object { $_.MessageSubject -like "*$Subject*" }
        }
        return $logs
    }
    catch { Write-Warning "Message tracking failed: $_"; return @() }
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — PROTOCOL LOG PARSING (SMTP Send/Receive)
# ═══════════════════════════════════════════════════════════════════════════════

function Parse-SmtpProtocolLog {
    <#
    .SYNOPSIS
        Парсит SMTP protocol log файлы Exchange (CSV-формат с #-комментариями).
        Поддерживает Send и Receive connector logs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [string]$Filter,
        [int]$MaxFiles = 10
    )

    if (-not (Test-Path $LogPath)) {
        Write-Warning "Path not found: $LogPath"
        return @()
    }

    $logFiles = Get-ChildItem -Path $LogPath -Filter '*.LOG' -Recurse |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First $MaxFiles

    if ($logFiles.Count -eq 0) { return @() }

    $allEntries = @()
    foreach ($file in $logFiles) {
        $content = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue
        $headers = $null

        foreach ($line in $content) {
            if ($line.StartsWith('#Fields:')) {
                $headers = ($line -replace '^#Fields:\s*', '') -split ','
                continue
            }
            if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) { continue }
            if (-not $headers) { continue }

            $values = $line -split ','
            $entry = [ordered]@{ SourceFile = $file.Name }
            for ($i = 0; $i -lt [Math]::Min($headers.Count, $values.Count); $i++) {
                $entry[$headers[$i].Trim()] = $values[$i].Trim()
            }
            $obj = [PSCustomObject]$entry
            if ($Filter) {
                $matched = $false
                foreach ($prop in $obj.PSObject.Properties) {
                    if ($prop.Value -and $prop.Value.ToString() -match [regex]::Escape($Filter)) {
                        $matched = $true; break
                    }
                }
                if (-not $matched) { continue }
            }
            $allEntries += $obj
        }
    }
    return $allEntries
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — TRANSPORT LOG SEARCH (text-based)
# ═══════════════════════════════════════════════════════════════════════════════

function Search-TransportLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][string]$Pattern,
        [int]$ContextLines = 3,
        [int]$MaxFiles = 50
    )

    if (-not (Test-Path $LogPath)) { Write-Warning "Path not found: $LogPath"; return @() }

    $logFiles = Get-ChildItem -Path $LogPath -Filter '*.log' -Recurse |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First $MaxFiles

    $allMatches = @()
    foreach ($file in $logFiles) {
        $lines = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match [regex]::Escape($Pattern)) {
                $s = [Math]::Max(0, $i - $ContextLines)
                $e = [Math]::Min($lines.Count - 1, $i + $ContextLines)
                $allMatches += [PSCustomObject]@{
                    File = $file.Name; FilePath = $file.FullName
                    Line = $i + 1; Match = $lines[$i]
                    Context = ($lines[$s..$e] -join "`r`n")
                    FileDate = $file.LastWriteTime
                }
            }
        }
    }
    return $allMatches
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — HEADER PARSING
# ═══════════════════════════════════════════════════════════════════════════════

function Parse-EmailHeaders {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$RawHeaders)

    $result = [PSCustomObject]@{
        Hops = @(); MessageId = $null; From = $null; To = $null
        Subject = $null; Date = $null; ContentType = $null
        SPF = $null; DKIM = $null; DMARC = $null
        XHeaders = @{}; TotalHops = 0; TotalDelayMs = 0
        ReturnPath = $null; AuthResults = $null
    }

    if ($RawHeaders -match '(?m)^Message-ID:\s*(.+)$')          { $result.MessageId = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^From:\s*(.+)$')                 { $result.From = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^To:\s*(.+)$')                   { $result.To = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^Subject:\s*(.+)$')              { $result.Subject = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^Date:\s*(.+)$')                 { $result.Date = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^Return-Path:\s*(.+)$')          { $result.ReturnPath = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^Content-Type:\s*(.+)$')         { $result.ContentType = $Matches[1].Trim() }
    if ($RawHeaders -match '(?m)^Authentication-Results:\s*(.+)$') { $result.AuthResults = $Matches[1].Trim() }

    if ($RawHeaders -match '(?m)spf=(\w+)')   { $result.SPF = $Matches[1] }
    if ($RawHeaders -match '(?m)dkim=(\w+)')  { $result.DKIM = $Matches[1] }
    if ($RawHeaders -match '(?m)dmarc=(\w+)') { $result.DMARC = $Matches[1] }

    $xMatches = [regex]::Matches($RawHeaders, '(?m)^(X-[\w-]+):\s*(.+)$')
    foreach ($m in $xMatches) { $result.XHeaders[$m.Groups[1].Value] = $m.Groups[2].Value.Trim() }

    $unfolded = $RawHeaders -replace '(\r?\n)\s+', ' '
    $rcvMatches = [regex]::Matches($unfolded, '(?m)^Received:\s*(.+?)(?=^[\w-]+:|\z)',
        [System.Text.RegularExpressions.RegexOptions]::Multiline)

    $hops = @()
    foreach ($m in $rcvMatches) {
        $line = $m.Groups[1].Value.Trim()
        $hop = [PSCustomObject]@{ Raw = $line; From = $null; By = $null; With = $null; For = $null; Timestamp = $null; Delay = $null; TLS = $false }
        if ($line -match 'from\s+([\w\.\-]+(?:\s*\([^\)]*\))?)') { $hop.From = $Matches[1].Trim() }
        if ($line -match 'by\s+([\w\.\-]+(?:\s*\([^\)]*\))?)')   { $hop.By = $Matches[1].Trim() }
        if ($line -match 'with\s+(\w+)')                          { $hop.With = $Matches[1] }
        if ($line -match 'for\s+<([^>]+)>')                       { $hop.For = $Matches[1] }
        if ($line -match 'TLS|STARTTLS|ESMTPS')                   { $hop.TLS = $true }
        if ($line -match ';\s*(.+)$') {
            try { $hop.Timestamp = [datetime]::Parse($Matches[1].Trim()) }
            catch { $hop.Timestamp = $Matches[1].Trim() }
        }
        $hops += $hop
    }

    [Array]::Reverse($hops)
    for ($i = 1; $i -lt $hops.Count; $i++) {
        if ($hops[$i].Timestamp -is [datetime] -and $hops[$i-1].Timestamp -is [datetime]) {
            $hops[$i].Delay = ($hops[$i].Timestamp - $hops[$i-1].Timestamp)
        }
    }

    $result.Hops = $hops
    $result.TotalHops = $hops.Count
    if ($hops.Count -ge 2 -and $hops[0].Timestamp -is [datetime] -and $hops[-1].Timestamp -is [datetime]) {
        $result.TotalDelayMs = ($hops[-1].Timestamp - $hops[0].Timestamp).TotalMilliseconds
    }
    return $result
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — TRANSPORT REPORTS (text output for GUI)
# ═══════════════════════════════════════════════════════════════════════════════

function Get-TransportReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Full','Queues','Connectors','AgentLog','RoutingTable','DSN','Summary','Pipeline','BackPressure')]
        [string]$ReportType,
        [string]$Server
    )

    $sp = @{}
    if ($Server) { $sp['Server'] = $Server }
    $L = [System.Collections.Generic.List[string]]::new()

    if ($ReportType -in @('Full','Queues','Summary')) {
        $L.Add("=== QUEUE STATUS ===`r`n")
        try {
            $q = Get-Queue @sp -ErrorAction Stop
            $grp = $q | Group-Object Status | Select-Object Name, Count
            $L.Add("Total queues: $($q.Count)")
            foreach ($g in $grp) { $L.Add("  $($g.Name): $($g.Count)") }
            $tot = ($q | Measure-Object -Property MessageCount -Sum).Sum
            $L.Add("Total messages in queues: $tot")
            $retry = $q | Where-Object { $_.Status -eq 'Retry' }
            if ($retry) {
                $L.Add(""); $L.Add("RETRY QUEUES:")
                foreach ($r in $retry) { $L.Add("  $($r.Identity) | $($r.MessageCount) msgs | $($r.LastError)") }
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -in @('Full','Connectors')) {
        $L.Add("=== SEND CONNECTORS ===`r`n")
        try {
            foreach ($c in (Get-SendConnector -ErrorAction Stop)) {
                $st = if ($c.Enabled) {'[ON]'} else {'[OFF]'}
                $L.Add("  $st $($c.Name) -> $($c.AddressSpaces -join ', ')")
                $L.Add("      SmartHosts: $($c.SmartHosts -join ', ')  |  MaxMsgSize: $($c.MaxMessageSize)")
                $L.Add("      TLS: $($c.RequireTLS)  |  Auth: $($c.SmartHostAuthMechanism)")
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add(""); $L.Add("=== RECEIVE CONNECTORS ===`r`n")
        try {
            foreach ($c in (Get-ReceiveConnector @sp -ErrorAction Stop)) {
                $st = if ($c.Enabled) {'[ON]'} else {'[OFF]'}
                $L.Add("  $st $($c.Name)")
                $L.Add("      Bindings: $($c.Bindings -join ', ')")
                $L.Add("      RemoteIPRanges: $($c.RemoteIPRanges -join ', ')")
                $L.Add("      Auth: $($c.AuthMechanism)  |  PermissionGroups: $($c.PermissionGroups)")
                $L.Add("      MaxMsgSize: $($c.MaxMessageSize)  |  MaxRecipients: $($c.MaxRecipientsPerMessage)")
                $L.Add("      TLS: $($c.RequireTLS)  |  ProtocolLogging: $($c.ProtocolLoggingLevel)")
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -in @('Full','AgentLog')) {
        $L.Add("=== TRANSPORT AGENTS ===`r`n")
        try {
            $agents = Get-TransportAgent -ErrorAction Stop
            foreach ($a in $agents) {
                $st = if ($a.Enabled) {'[ON]'} else {'[OFF]'}
                $L.Add("  $st Priority:$($a.Priority) $($a.Identity)")
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add(""); $L.Add("=== AGENT LOG (last 100) ===`r`n")
        try {
            $al = Get-AgentLog @sp -ErrorAction Stop | Select-Object -Last 100 Timestamp, Agent, Event, Action, SmtpResponse, P1FromAddress, IPAddress, MessageId
            foreach ($e in $al) {
                $L.Add("  $($e.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')) | $($e.Agent) | $($e.Event) | $($e.Action)")
                if ($e.SmtpResponse) { $L.Add("    Response: $($e.SmtpResponse)") }
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -in @('Full','RoutingTable')) {
        $L.Add("=== TRANSPORT CONFIGURATION ===`r`n")
        try {
            $cfg = Get-TransportConfig -ErrorAction Stop
            $L.Add("  Max receive size       : $($cfg.MaxReceiveSize)")
            $L.Add("  Max send size          : $($cfg.MaxSendSize)")
            $L.Add("  Max recipients         : $($cfg.MaxRecipientEnvelopeLimit)")
            $L.Add("  Shadow redundancy      : $($cfg.ShadowRedundancyEnabled)")
            $L.Add("  Safety net hold time   : $($cfg.SafetyNetHoldTime)")
            $L.Add("  Max header size        : $($cfg.MaxHeaderSize)")
            $L.Add("  TLS send domain secure : $($cfg.TLSSendDomainSecureList -join ', ')")
            $L.Add("  TLS recv domain secure : $($cfg.TLSReceiveDomainSecureList -join ', ')")
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
        $L.Add("=== TRANSPORT SERVICE ===`r`n")
        try {
            $ts = Get-TransportService @sp -ErrorAction Stop
            $L.Add("  Message tracking enabled : $($ts.MessageTrackingLogEnabled)")
            $L.Add("  Message tracking path    : $($ts.MessageTrackingLogPath)")
            $L.Add("  Max tracking log age     : $($ts.MessageTrackingLogMaxAge)")
            $L.Add("  Connectivity log enabled : $($ts.ConnectivityLogEnabled)")
            $L.Add("  Connectivity log path    : $($ts.ConnectivityLogPath)")
            $L.Add("  Send protocol log path   : $($ts.SendProtocolLogPath)")
            $L.Add("  Recv protocol log path   : $($ts.ReceiveProtocolLogPath)")
            $L.Add("  Routing log path         : $($ts.RoutingTableLogPath)")
            $L.Add("  Pipeline tracing enabled : $($ts.PipelineTracingEnabled)")
            $L.Add("  Pipeline tracing path    : $($ts.PipelineTracingPath)")
            $L.Add("  Content conversion path  : $($ts.ContentConversionTracingEnabled)")
            $L.Add("  Pickup directory         : $($ts.PickupDirectoryPath)")
            $L.Add("  Replay directory         : $($ts.ReplayDirectoryPath)")
            $L.Add("  Max outbound connections : $($ts.MaxOutboundConnections)")
            $L.Add("  Max per-domain out conn  : $($ts.MaxPerDomainOutboundConnections)")
            $L.Add("  Transient failure retry  : $($ts.TransientFailureRetryCount) x $($ts.TransientFailureRetryInterval)")
            $L.Add("  Outbound conn failure    : $($ts.OutboundConnectionFailureRetryInterval)")
            $L.Add("  Message retry interval   : $($ts.MessageRetryInterval)")
            $L.Add("  Message expiration       : $($ts.MessageExpirationTimeout)")
            $L.Add("  Delay DSN timeout        : $($ts.DelayNotificationTimeout)")
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -in @('Full','DSN')) {
        $L.Add("=== DSN CONFIGURATION ===`r`n")
        try {
            $cfg = Get-TransportConfig -ErrorAction Stop
            $L.Add("  Generate DSN copy for      : $($cfg.GenerateCopyOfDSNFor -join ', ')")
            $L.Add("  External postmaster        : $($cfg.ExternalPostmasterAddress)")
            $L.Add("  External DSN max msg size  : $($cfg.ExternalDsnMaxMessageAttachSize)")
            $L.Add("  Internal DSN max msg size  : $($cfg.InternalDsnMaxMessageAttachSize)")
            $L.Add("")
            $L.Add("  Recent DSN events (24h):")
            $dp = @{ EventId = 'DSN'; Start = (Get-Date).AddDays(-1); End = Get-Date; ResultSize = 100 }
            if ($Server) { $dp['Server'] = $Server }
            $dsn = Get-MessageTrackingLog @dp -ErrorAction Stop
            if ($dsn.Count -eq 0) { $L.Add("    No DSN events") }
            else {
                foreach ($d in $dsn) {
                    $L.Add("    $($d.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')) | $($d.Sender) -> $($d.Recipients -join ', ')")
                    if ($d.RecipientStatus) { $L.Add("      Status: $($d.RecipientStatus)") }
                }
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -eq 'Summary') {
        $L.Add("=== DELIVERY SUMMARY (last 24h) ===`r`n")
        try {
            $tp = @{ Start = (Get-Date).AddDays(-1); End = Get-Date; ResultSize = 'Unlimited' }
            if ($Server) { $tp['Server'] = $Server }
            $all = Get-MessageTrackingLog @tp -ErrorAction Stop
            $grp = $all | Group-Object EventId | Select-Object Name, Count | Sort-Object Count -Descending
            foreach ($s in $grp) { $L.Add("    $($s.Name.PadRight(15)) $($s.Count)") }
            $del = ($all | Where-Object EventId -eq 'DELIVER').Count
            $fail = ($all | Where-Object EventId -eq 'FAIL').Count
            $defer = ($all | Where-Object EventId -eq 'DEFER').Count
            $t = $del + $fail
            if ($t -gt 0) {
                $rate = [math]::Round(($del / $t) * 100, 1)
                $L.Add(""); $L.Add("    Delivery rate: $rate% ($del delivered / $fail failed)")
                $L.Add("    Deferred: $defer")
            }
            $L.Add("")
            $L.Add("  By Server:")
            $srvGrp = $all | Group-Object ServerHostname | Sort-Object Count -Descending
            foreach ($s in $srvGrp) { $L.Add("    $($s.Name.PadRight(30)) $($s.Count) events") }
            $L.Add("")
            $L.Add("  Top Senders (by volume):")
            $sndGrp = $all | Where-Object EventId -eq 'RECEIVE' | Group-Object Sender | Sort-Object Count -Descending | Select-Object -First 20
            foreach ($s in $sndGrp) { $L.Add("    $($s.Count.ToString().PadLeft(6)) | $($s.Name)") }
            $L.Add("")
            $L.Add("  Recent Failures:")
            $failures = $all | Where-Object EventId -eq 'FAIL' | Select-Object -Last 20
            foreach ($f in $failures) {
                $L.Add("    $($f.Timestamp.ToString('HH:mm:ss')) $($f.Sender) -> $($f.Recipients -join ', ')")
                if ($f.RecipientStatus) { $L.Add("      $($f.RecipientStatus)") }
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    if ($ReportType -eq 'Pipeline') {
        $L.Add("=== TRANSPORT PIPELINE ===`r`n")
        try {
            $agents = Get-TransportAgent -ErrorAction Stop | Sort-Object Priority
            $L.Add("  Registered Transport Agents (by priority):`r`n")
            foreach ($a in $agents) {
                $st = if ($a.Enabled) {'ENABLED'} else {'DISABLED'}
                $L.Add("  [$($a.Priority.ToString().PadLeft(3))] $($a.Identity)")
                $L.Add("        Status: $st  |  Assembly: $($a.AssemblyPath)")
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
        try {
            $ts = Get-TransportService @sp -ErrorAction Stop
            $L.Add("  Pipeline Tracing: $($ts.PipelineTracingEnabled)")
            $L.Add("  Pipeline Path   : $($ts.PipelineTracingPath)")
            $L.Add("  Pipeline Sender : $($ts.PipelineTracingSenderAddress)")
        } catch { $L.Add("(could not get transport service info)") }
        $L.Add("")
    }

    if ($ReportType -eq 'BackPressure') {
        $L.Add("=== BACK PRESSURE STATUS ===`r`n")
        try {
            $diag = Get-ExchangeDiagnosticInfo -Process EdgeTransport -Component ResourceThrottling -ErrorAction Stop
            $L.Add($diag)
        } catch { $L.Add("(Get-ExchangeDiagnosticInfo not available or failed: $_)") }
        $L.Add("")
        try {
            $ts = Get-TransportService @sp -ErrorAction Stop
            $L.Add("  Resource monitoring:")
            foreach ($prop in @('DatabaseMaxCacheSize','DatabasePath','QueueDatabasePath',
                                'QueueDatabaseMaxCacheSize','QueueDatabaseLoggingPath')) {
                try { $L.Add("    $prop : $($ts.$prop)") } catch {}
            }
        } catch { $L.Add("ERROR: $_") }
        $L.Add("")
    }

    return ($L -join "`r`n")
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — DASHBOARD DATA
# ═══════════════════════════════════════════════════════════════════════════════

function Get-DashboardData {
    [CmdletBinding()]
    param([string]$Server)

    $sp = @{}
    if ($Server) { $sp['Server'] = $Server }
    $L = [System.Collections.Generic.List[string]]::new()

    $L.Add("  EXCHANGE TRANSPORT DASHBOARD")
    $L.Add("  Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $L.Add("  Server: $($Server ?? 'All')")
    $L.Add("  " + ("=" * 60))

    # Queues
    $L.Add(""); $L.Add("  QUEUES")
    $L.Add("  " + ("-" * 40))
    try {
        $q = Get-Queue @sp -ErrorAction Stop
        $total = ($q | Measure-Object -Property MessageCount -Sum).Sum
        $ready = ($q | Where-Object Status -eq 'Ready' | Measure-Object -Property MessageCount -Sum).Sum
        $retry = ($q | Where-Object Status -eq 'Retry' | Measure-Object -Property MessageCount -Sum).Sum
        $active = ($q | Where-Object Status -eq 'Active' | Measure-Object -Property MessageCount -Sum).Sum
        $suspended = ($q | Where-Object Status -eq 'Suspended' | Measure-Object -Property MessageCount -Sum).Sum

        $L.Add("  Total queues      : $($q.Count)")
        $L.Add("  Total messages    : $total")
        $L.Add("    Ready           : $ready")
        $L.Add("    Active          : $active")
        $L.Add("    Retry           : $retry $(if ($retry -gt 0) {'  *** ATTENTION ***'})")
        $L.Add("    Suspended       : $suspended $(if ($suspended -gt 0) {'  *** CHECK ***'})")

        $retryQ = $q | Where-Object { $_.Status -eq 'Retry' -and $_.MessageCount -gt 0 }
        if ($retryQ) {
            $L.Add(""); $L.Add("    Retry queue details:")
            foreach ($r in $retryQ) {
                $L.Add("      $($r.Identity)")
                $L.Add("        Messages: $($r.MessageCount)  |  NextRetry: $($r.NextRetryTime)")
                $L.Add("        Error: $($r.LastError)")
            }
        }
    } catch { $L.Add("  ERROR: $_") }

    # Delivery stats (1h and 24h)
    foreach ($hours in @(1, 24)) {
        $L.Add(""); $L.Add("  DELIVERY ($($hours)h)")
        $L.Add("  " + ("-" * 40))
        try {
            $tp = @{ Start = (Get-Date).AddHours(-$hours); End = Get-Date; ResultSize = 'Unlimited' }
            if ($Server) { $tp['Server'] = $Server }
            $all = Get-MessageTrackingLog @tp -ErrorAction Stop

            $receive = ($all | Where-Object EventId -eq 'RECEIVE').Count
            $deliver = ($all | Where-Object EventId -eq 'DELIVER').Count
            $send = ($all | Where-Object EventId -eq 'SEND').Count
            $fail = ($all | Where-Object EventId -eq 'FAIL').Count
            $defer = ($all | Where-Object EventId -eq 'DEFER').Count
            $dsn = ($all | Where-Object EventId -eq 'DSN').Count

            $t = $deliver + $fail
            $rate = if ($t -gt 0) { [math]::Round(($deliver / $t) * 100, 1) } else { 'N/A' }

            $L.Add("  Received    : $receive")
            $L.Add("  Delivered   : $deliver")
            $L.Add("  Sent (relay): $send")
            $L.Add("  Failed      : $fail $(if ($fail -gt 0) {'  !!!'})")
            $L.Add("  Deferred    : $defer $(if ($defer -gt 10) {'  *** HIGH ***'})")
            $L.Add("  DSN/NDR     : $dsn")
            $L.Add("  Delivery %  : $rate%")
        } catch { $L.Add("  ERROR: $_") }
    }

    # Connectors health
    $L.Add(""); $L.Add("  CONNECTORS")
    $L.Add("  " + ("-" * 40))
    try {
        $sc = Get-SendConnector -ErrorAction Stop
        $scOn = ($sc | Where-Object Enabled).Count
        $scOff = ($sc | Where-Object { -not $_.Enabled }).Count
        $L.Add("  Send connectors    : $scOn active, $scOff disabled")
    } catch { $L.Add("  Send connectors: ERROR") }
    try {
        $rc = Get-ReceiveConnector @sp -ErrorAction Stop
        $rcOn = ($rc | Where-Object Enabled).Count
        $rcOff = ($rc | Where-Object { -not $_.Enabled }).Count
        $L.Add("  Receive connectors : $rcOn active, $rcOff disabled")
    } catch { $L.Add("  Receive connectors: ERROR") }

    $L.Add(""); $L.Add("  " + ("=" * 60))
    return ($L -join "`r`n")
}

# ═══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS — EXPORT
# ═══════════════════════════════════════════════════════════════════════════════

function Export-ResultsToFile {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Data, [Parameter(Mandatory)][string]$FilePath, [string]$Format = 'CSV')
    switch ($Format) {
        'CSV'  { $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 }
        'JSON' { $Data | ConvertTo-Json -Depth 5 | Set-Content -Path $FilePath -Encoding UTF8 }
        default { $Data | Out-File -FilePath $FilePath -Encoding UTF8 }
    }
}

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
    $dgv.MultiSelect = $Multi.IsPresent
    $dgv.AutoSizeColumnsMode = 'Fill'
    $dgv.RowHeadersVisible = $false
    $dgv.BackgroundColor = [System.Drawing.Color]::White
    $dgv.BorderStyle = 'None'
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 250)
    return $dgv
}

function New-Btn {
    param([string]$Text, [int]$W = 110, [string]$Color = '')
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Size = New-Object System.Drawing.Size($W, 28)
    $b.FlatStyle = 'Flat'; $b.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
    switch ($Color) {
        'Blue'  { $b.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215);  $b.ForeColor = [System.Drawing.Color]::White }
        'Red'   { $b.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50);  $b.ForeColor = [System.Drawing.Color]::White }
        'Green' { $b.BackColor = [System.Drawing.Color]::FromArgb(50, 150, 50);  $b.ForeColor = [System.Drawing.Color]::White }
        'Orange'{ $b.BackColor = [System.Drawing.Color]::FromArgb(200, 130, 0);  $b.ForeColor = [System.Drawing.Color]::White }
    }
    return $b
}

function New-ConsoleTextBox {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline = $true; $t.ReadOnly = $true
    $t.ScrollBars = 'Both'; $t.WordWrap = $false
    $t.Font = New-Object System.Drawing.Font('Consolas', 10)
    $t.BackColor = [System.Drawing.Color]::FromArgb(25, 25, 35)
    $t.ForeColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    return $t
}

function New-FlowBar {
    param([int]$H = 40)
    $f = New-Object System.Windows.Forms.FlowLayoutPanel
    $f.Dock = 'Top'; $f.Height = $H
    $f.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 0)
    return $f
}

function New-BoldLabel {
    param([string]$Text)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Dock = 'Top'; $l.AutoSize = $true
    $l.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $l.Padding = New-Object System.Windows.Forms.Padding(5, 5, 0, 2)
    return $l
}

function New-InlineLabel {
    param([string]$Text, [int]$MarginLeft = 0)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding($MarginLeft, 6, 5, 0)
    return $l
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN GUI
# ═══════════════════════════════════════════════════════════════════════════════

function Show-ExchangeRetryGUI {
    [CmdletBinding()]
    param()

    $script:ExSession = $null
    $script:LastParsedHeaders = $null
    $script:LastTrackingResults = $null
    $script:LastLogResults = $null
    $script:LastProtoResults = $null

    # ── Form ──
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'ExchangeRetry v0.3 — Exchange Transport Manager'
    $form.Size = New-Object System.Drawing.Size(1300, 850)
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 600)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    # ── Top Panel ──
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'; $panelTop.Height = 50

    $lblSrv = New-Object System.Windows.Forms.Label
    $lblSrv.Text = 'Exchange Server:'; $lblSrv.Location = New-Object System.Drawing.Point(10, 15); $lblSrv.AutoSize = $true

    $txtServer = New-Object System.Windows.Forms.TextBox
    $txtServer.Location = New-Object System.Drawing.Point(130, 12); $txtServer.Size = New-Object System.Drawing.Size(200, 23)
    $txtServer.Text = $script:Config.ExchangeServer

    $btnConnect = New-Btn -Text 'Connect' -W 90 -Color 'Blue'
    $btnConnect.Location = New-Object System.Drawing.Point(340, 10)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Disconnected'; $lblStatus.ForeColor = [System.Drawing.Color]::Gray
    $lblStatus.Location = New-Object System.Drawing.Point(440, 15); $lblStatus.AutoSize = $true

    # Scope: select which server to query
    $lblScope = New-Object System.Windows.Forms.Label
    $lblScope.Text = 'Scope:'; $lblScope.Location = New-Object System.Drawing.Point(650, 15); $lblScope.AutoSize = $true

    $cmbScope = New-Object System.Windows.Forms.ComboBox
    $cmbScope.DropDownStyle = 'DropDownList'
    $cmbScope.Location = New-Object System.Drawing.Point(695, 12); $cmbScope.Size = New-Object System.Drawing.Size(250, 23)
    $cmbScope.Items.Add('(All Servers)') | Out-Null
    $cmbScope.SelectedIndex = 0

    $script:ExchangeServers = @()

    $panelTop.Controls.AddRange(@($lblSrv, $txtServer, $btnConnect, $lblStatus, $lblScope, $cmbScope))

    # ── Status Bar ──
    $statusBar = New-Object System.Windows.Forms.StatusStrip
    $stLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $stLabel.Text = 'Ready'
    $statusBar.Items.Add($stLabel) | Out-Null

    # ── Tab Control ──
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'

    # ══════════════════════════════════════════════════════════════════════
    # TAB 0: DASHBOARD
    # ══════════════════════════════════════════════════════════════════════
    $tabDash = New-Object System.Windows.Forms.TabPage; $tabDash.Text = 'Dashboard'; $tabDash.Padding = New-Object System.Windows.Forms.Padding(5)

    $barDash = New-FlowBar
    $btnRefreshDash = New-Btn -Text 'Refresh Dashboard' -W 140 -Color 'Blue'
    $barDash.Controls.Add($btnRefreshDash)

    $txtDash = New-ConsoleTextBox; $txtDash.Dock = 'Fill'
    $txtDash.Font = New-Object System.Drawing.Font('Consolas', 11)
    $tabDash.Controls.Add($txtDash); $tabDash.Controls.Add($barDash)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 1: QUEUES
    # ══════════════════════════════════════════════════════════════════════
    $tabQ = New-Object System.Windows.Forms.TabPage; $tabQ.Text = 'Queues'; $tabQ.Padding = New-Object System.Windows.Forms.Padding(5)

    # Main vertical split: queues+messages (top) | errors (bottom)
    $splitQMain = New-Object System.Windows.Forms.SplitContainer
    $splitQMain.Dock = 'Fill'; $splitQMain.Orientation = 'Horizontal'; $splitQMain.SplitterDistance = 450
    $splitQMain.Panel2MinSize = 100

    $splitQ = New-Object System.Windows.Forms.SplitContainer
    $splitQ.Dock = 'Fill'; $splitQ.Orientation = 'Horizontal'; $splitQ.SplitterDistance = 220

    # Top: queues
    $lblQ = New-BoldLabel 'Queues'
    $dgvQ = New-StyledDGV; $dgvQ.Dock = 'Fill'
    $barQ = New-FlowBar
    $btnRefQ = New-Btn -Text 'Refresh' -W 90
    $btnRetryQ = New-Btn -Text 'Retry Queue' -W 100 -Color 'Blue'
    $txtQFilter = New-Object System.Windows.Forms.TextBox
    $txtQFilter.Size = New-Object System.Drawing.Size(250, 23)
    $txtQFilter.PlaceholderText = 'Filter: Status -eq "Retry"'
    $txtQFilter.Margin = New-Object System.Windows.Forms.Padding(10, 3, 0, 0)
    $chkAutoRef = New-Object System.Windows.Forms.CheckBox
    $chkAutoRef.Text = "Auto ($($script:Config.RefreshIntervalSec)s)"; $chkAutoRef.AutoSize = $true
    $chkAutoRef.Margin = New-Object System.Windows.Forms.Padding(10, 5, 0, 0)
    $barQ.Controls.AddRange(@($btnRefQ, $btnRetryQ, $txtQFilter, $chkAutoRef))
    $splitQ.Panel1.Controls.Add($dgvQ); $splitQ.Panel1.Controls.Add($barQ); $splitQ.Panel1.Controls.Add($lblQ)

    # Middle: messages
    $lblMsg = New-BoldLabel 'Messages'
    $dgvMsg = New-StyledDGV -Multi; $dgvMsg.Dock = 'Fill'
    $barMsg = New-FlowBar
    $btnRetryM = New-Btn -Text 'Retry Selected' -Color 'Blue'
    $btnSuspM = New-Btn -Text 'Suspend Selected' -W 130
    $btnRemM = New-Btn -Text 'Remove Selected' -W 120 -Color 'Red'
    $chkNDR = New-Object System.Windows.Forms.CheckBox; $chkNDR.Text = 'NDR'; $chkNDR.AutoSize = $true
    $chkNDR.Margin = New-Object System.Windows.Forms.Padding(10, 6, 0, 0)
    $lblMsgCnt = New-Object System.Windows.Forms.Label; $lblMsgCnt.AutoSize = $true
    $lblMsgCnt.Margin = New-Object System.Windows.Forms.Padding(20, 8, 0, 0); $lblMsgCnt.ForeColor = [System.Drawing.Color]::Gray
    $barMsg.Controls.AddRange(@($btnRetryM, $btnSuspM, $btnRemM, $chkNDR, $lblMsgCnt))
    $splitQ.Panel2.Controls.Add($dgvMsg); $splitQ.Panel2.Controls.Add($barMsg); $splitQ.Panel2.Controls.Add($lblMsg)

    $splitQMain.Panel1.Controls.Add($splitQ)

    # Bottom: recent errors panel
    $lblErrors = New-BoldLabel 'Recent Errors (FAIL/DEFER/DSN — last 24h)'
    $lblErrors.ForeColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $dgvErrors = New-StyledDGV -Multi; $dgvErrors.Dock = 'Fill'
    $dgvErrors.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 245, 245)
    $barErrors = New-FlowBar; $barErrors.Dock = 'Bottom'
    $btnRefErrors = New-Btn -Text 'Refresh Errors' -W 120
    $btnExportErrors = New-Btn -Text 'Export...' -W 90 -Color 'Green'
    $lblErrCnt = New-Object System.Windows.Forms.Label; $lblErrCnt.AutoSize = $true
    $lblErrCnt.ForeColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
    $lblErrCnt.Margin = New-Object System.Windows.Forms.Padding(20, 8, 0, 0)
    $barErrors.Controls.AddRange(@($btnRefErrors, $btnExportErrors, $lblErrCnt))
    $splitQMain.Panel2.Controls.Add($dgvErrors); $splitQMain.Panel2.Controls.Add($barErrors); $splitQMain.Panel2.Controls.Add($lblErrors)

    $tabQ.Controls.Add($splitQMain)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 2: MESSAGE TRACKING
    # ══════════════════════════════════════════════════════════════════════
    $tabTrack = New-Object System.Windows.Forms.TabPage; $tabTrack.Text = 'Message Tracking'; $tabTrack.Padding = New-Object System.Windows.Forms.Padding(5)

    # Search bar row 1
    $barTr1 = New-FlowBar -H 35
    $barTr1.Controls.AddRange(@(
        (New-InlineLabel 'Message-ID:'),
        ($txtTrMsgId = New-Object System.Windows.Forms.TextBox),
        (New-InlineLabel 'Sender:' -MarginLeft 10),
        ($txtTrSender = New-Object System.Windows.Forms.TextBox),
        (New-InlineLabel 'Recipient:' -MarginLeft 10),
        ($txtTrRecip = New-Object System.Windows.Forms.TextBox)
    ))
    $txtTrMsgId.Size = New-Object System.Drawing.Size(200, 23)
    $txtTrSender.Size = New-Object System.Drawing.Size(180, 23)
    $txtTrRecip.Size = New-Object System.Drawing.Size(180, 23)

    # Search bar row 2
    $barTr2 = New-FlowBar -H 38
    $lblTrStart = New-InlineLabel 'From:'
    $dtpTrStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpTrStart.Format = 'Custom'; $dtpTrStart.CustomFormat = 'yyyy-MM-dd HH:mm'; $dtpTrStart.Value = (Get-Date).AddDays(-1)
    $dtpTrStart.Size = New-Object System.Drawing.Size(150, 23)

    $lblTrEnd = New-InlineLabel 'To:' -MarginLeft 10
    $dtpTrEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpTrEnd.Format = 'Custom'; $dtpTrEnd.CustomFormat = 'yyyy-MM-dd HH:mm'; $dtpTrEnd.Value = Get-Date
    $dtpTrEnd.Size = New-Object System.Drawing.Size(150, 23)

    $lblTrEvent = New-InlineLabel 'EventId:' -MarginLeft 10
    $cmbTrEvent = New-Object System.Windows.Forms.ComboBox
    $cmbTrEvent.DropDownStyle = 'DropDownList'; $cmbTrEvent.Size = New-Object System.Drawing.Size(110, 23)
    $cmbTrEvent.Items.AddRange(@('(All)','RECEIVE','SEND','DELIVER','SUBMIT','FAIL','DSN','DEFER','EXPAND','REDIRECT','RESOLVE','TRANSFER','POISONMESSAGE'))
    $cmbTrEvent.SelectedIndex = 0

    $lblTrSubject = New-InlineLabel 'Subject:' -MarginLeft 10
    $txtTrSubject = New-Object System.Windows.Forms.TextBox
    $txtTrSubject.Size = New-Object System.Drawing.Size(150, 23)

    $btnTrSearch = New-Btn -Text 'Search' -W 90 -Color 'Blue'

    $barTr2.Controls.AddRange(@($lblTrStart, $dtpTrStart, $lblTrEnd, $dtpTrEnd, $lblTrEvent, $cmbTrEvent, $lblTrSubject, $txtTrSubject, $btnTrSearch))

    # Results
    $dgvTrack = New-StyledDGV -Multi; $dgvTrack.Dock = 'Fill'
    $barTrBot = New-FlowBar; $barTrBot.Dock = 'Bottom'
    $btnTrExport = New-Btn -Text 'Export...' -W 90 -Color 'Green'
    $btnTrPath = New-Btn -Text 'Show Message Path' -W 140 -Color 'Orange'
    $lblTrCnt = New-Object System.Windows.Forms.Label; $lblTrCnt.AutoSize = $true
    $lblTrCnt.ForeColor = [System.Drawing.Color]::Gray; $lblTrCnt.Margin = New-Object System.Windows.Forms.Padding(20, 8, 0, 0)
    $barTrBot.Controls.AddRange(@($btnTrExport, $btnTrPath, $lblTrCnt))

    $tabTrack.Controls.Add($dgvTrack); $tabTrack.Controls.Add($barTrBot); $tabTrack.Controls.Add($barTr2); $tabTrack.Controls.Add($barTr1)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 3: PROTOCOL LOGS (SMTP Send/Receive)
    # ══════════════════════════════════════════════════════════════════════
    $tabProto = New-Object System.Windows.Forms.TabPage; $tabProto.Text = 'Protocol Logs'; $tabProto.Padding = New-Object System.Windows.Forms.Padding(5)

    $barProto = New-FlowBar -H 40
    $lblProtoPath = New-InlineLabel 'Log Path:'
    $txtProtoPath = New-Object System.Windows.Forms.TextBox
    $txtProtoPath.Size = New-Object System.Drawing.Size(350, 23)
    $txtProtoPath.PlaceholderText = '\\server\SendProtocolLog or ReceiveProtocolLog'
    $btnProtoBrowse = New-Btn -Text 'Browse...' -W 80
    $lblProtoFilter = New-InlineLabel 'Filter:' -MarginLeft 10
    $txtProtoFilter = New-Object System.Windows.Forms.TextBox
    $txtProtoFilter.Size = New-Object System.Drawing.Size(200, 23)
    $txtProtoFilter.PlaceholderText = 'user@domain.com'
    $lblProtoMax = New-InlineLabel 'Files:' -MarginLeft 10
    $nudProtoMax = New-Object System.Windows.Forms.NumericUpDown
    $nudProtoMax.Minimum = 1; $nudProtoMax.Maximum = 100; $nudProtoMax.Value = 10
    $nudProtoMax.Size = New-Object System.Drawing.Size(55, 23)
    $btnProtoLoad = New-Btn -Text 'Load & Parse' -W 110 -Color 'Blue'

    $barProto.Controls.AddRange(@($lblProtoPath, $txtProtoPath, $btnProtoBrowse, $lblProtoFilter, $txtProtoFilter, $lblProtoMax, $nudProtoMax, $btnProtoLoad))

    $dgvProto = New-StyledDGV -Multi; $dgvProto.Dock = 'Fill'
    $barProtoBot = New-FlowBar; $barProtoBot.Dock = 'Bottom'
    $btnProtoExport = New-Btn -Text 'Export...' -W 90 -Color 'Green'
    $lblProtoCnt = New-Object System.Windows.Forms.Label; $lblProtoCnt.AutoSize = $true
    $lblProtoCnt.ForeColor = [System.Drawing.Color]::Gray; $lblProtoCnt.Margin = New-Object System.Windows.Forms.Padding(20, 8, 0, 0)
    $barProtoBot.Controls.AddRange(@($btnProtoExport, $lblProtoCnt))

    $tabProto.Controls.Add($dgvProto); $tabProto.Controls.Add($barProtoBot); $tabProto.Controls.Add($barProto)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 4: TRANSPORT LOG SEARCH
    # ══════════════════════════════════════════════════════════════════════
    $tabLogs = New-Object System.Windows.Forms.TabPage; $tabLogs.Text = 'Log Search'; $tabLogs.Padding = New-Object System.Windows.Forms.Padding(5)

    $barLog = New-FlowBar -H 40
    $lblLogPath = New-InlineLabel 'Path:'
    $txtLogPath = New-Object System.Windows.Forms.TextBox
    $txtLogPath.Size = New-Object System.Drawing.Size(300, 23)
    $txtLogPath.PlaceholderText = '\\server\TransportLogs'
    $btnLogBrowse = New-Btn -Text 'Browse...' -W 80
    $lblLogPat = New-InlineLabel 'Pattern:' -MarginLeft 10
    $txtLogPat = New-Object System.Windows.Forms.TextBox
    $txtLogPat.Size = New-Object System.Drawing.Size(200, 23)
    $btnLogSearch = New-Btn -Text 'Search' -W 90 -Color 'Blue'
    $barLog.Controls.AddRange(@($lblLogPath, $txtLogPath, $btnLogBrowse, $lblLogPat, $txtLogPat, $btnLogSearch))

    $splitLog = New-Object System.Windows.Forms.SplitContainer
    $splitLog.Dock = 'Fill'; $splitLog.Orientation = 'Horizontal'; $splitLog.SplitterDistance = 250

    $dgvLog = New-StyledDGV; $dgvLog.Dock = 'Fill'
    $lblLogCnt = New-BoldLabel 'Results'
    $splitLog.Panel1.Controls.Add($dgvLog); $splitLog.Panel1.Controls.Add($lblLogCnt)

    $lblLogCtx = New-BoldLabel 'Context'
    $txtLogCtx = New-ConsoleTextBox; $txtLogCtx.Dock = 'Fill'
    $barLogBot = New-FlowBar; $barLogBot.Dock = 'Bottom'
    $btnLogExport = New-Btn -Text 'Export...' -W 90 -Color 'Green'
    $barLogBot.Controls.Add($btnLogExport)
    $splitLog.Panel2.Controls.Add($txtLogCtx); $splitLog.Panel2.Controls.Add($barLogBot); $splitLog.Panel2.Controls.Add($lblLogCtx)

    $tabLogs.Controls.Add($splitLog); $tabLogs.Controls.Add($barLog)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 5: HEADER ANALYZER
    # ══════════════════════════════════════════════════════════════════════
    $tabHdr = New-Object System.Windows.Forms.TabPage; $tabHdr.Text = 'Header Analyzer'; $tabHdr.Padding = New-Object System.Windows.Forms.Padding(5)

    $splitHdr = New-Object System.Windows.Forms.SplitContainer
    $splitHdr.Dock = 'Fill'; $splitHdr.Orientation = 'Horizontal'; $splitHdr.SplitterDistance = 180

    $lblHdrIn = New-BoldLabel 'Paste email headers:'
    $txtHdrIn = New-Object System.Windows.Forms.TextBox
    $txtHdrIn.Dock = 'Fill'; $txtHdrIn.Multiline = $true; $txtHdrIn.ScrollBars = 'Both'; $txtHdrIn.WordWrap = $false
    $txtHdrIn.Font = New-Object System.Drawing.Font('Consolas', 9)
    $barHdr = New-FlowBar; $barHdr.Dock = 'Bottom'
    $btnHdrParse = New-Btn -Text 'Analyze' -W 100 -Color 'Blue'
    $btnHdrFile = New-Btn -Text 'Load File...' -W 100
    $btnHdrExport = New-Btn -Text 'Export...' -W 90 -Color 'Green'
    $barHdr.Controls.AddRange(@($btnHdrParse, $btnHdrFile, $btnHdrExport))
    $splitHdr.Panel1.Controls.Add($txtHdrIn); $splitHdr.Panel1.Controls.Add($barHdr); $splitHdr.Panel1.Controls.Add($lblHdrIn)

    # Result area: info + hops | X-headers
    $lblHdrRes = New-BoldLabel 'Results'
    $splitHdrRes = New-Object System.Windows.Forms.SplitContainer
    $splitHdrRes.Dock = 'Fill'; $splitHdrRes.SplitterDistance = 550

    $txtHdrInfo = New-Object System.Windows.Forms.TextBox
    $txtHdrInfo.Dock = 'Top'; $txtHdrInfo.Multiline = $true; $txtHdrInfo.ReadOnly = $true; $txtHdrInfo.Height = 110
    $txtHdrInfo.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtHdrInfo.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 255); $txtHdrInfo.ScrollBars = 'Vertical'
    $dgvHops = New-StyledDGV; $dgvHops.Dock = 'Fill'
    $splitHdrRes.Panel1.Controls.Add($dgvHops); $splitHdrRes.Panel1.Controls.Add($txtHdrInfo)

    $lblXH = New-BoldLabel 'X-Headers'
    $dgvXH = New-StyledDGV; $dgvXH.Dock = 'Fill'
    $splitHdrRes.Panel2.Controls.Add($dgvXH); $splitHdrRes.Panel2.Controls.Add($lblXH)

    $splitHdr.Panel2.Controls.Add($splitHdrRes); $splitHdr.Panel2.Controls.Add($lblHdrRes)
    $tabHdr.Controls.Add($splitHdr)

    # ══════════════════════════════════════════════════════════════════════
    # TAB 6: REPORTS
    # ══════════════════════════════════════════════════════════════════════
    $tabRpt = New-Object System.Windows.Forms.TabPage; $tabRpt.Text = 'Reports'; $tabRpt.Padding = New-Object System.Windows.Forms.Padding(5)

    $barRpt = New-FlowBar -H 40
    $lblRptType = New-InlineLabel 'Report:'
    $cmbRpt = New-Object System.Windows.Forms.ComboBox
    $cmbRpt.DropDownStyle = 'DropDownList'; $cmbRpt.Size = New-Object System.Drawing.Size(150, 23)
    $cmbRpt.Items.AddRange(@('Full','Queues','Connectors','AgentLog','RoutingTable','DSN','Summary','Pipeline','BackPressure'))
    $cmbRpt.SelectedIndex = 0
    $btnRunRpt = New-Btn -Text 'Run Report' -W 110 -Color 'Blue'
    $btnSaveRpt = New-Btn -Text 'Save to File...' -W 120 -Color 'Green'
    $barRpt.Controls.AddRange(@($lblRptType, $cmbRpt, $btnRunRpt, $btnSaveRpt))

    $txtRpt = New-ConsoleTextBox; $txtRpt.Dock = 'Fill'
    $txtRpt.Font = New-Object System.Drawing.Font('Consolas', 10)
    $tabRpt.Controls.Add($txtRpt); $tabRpt.Controls.Add($barRpt)

    # ── Add all tabs ──
    $tabs.TabPages.AddRange(@($tabDash, $tabQ, $tabTrack, $tabProto, $tabLogs, $tabHdr, $tabRpt))

    # ── Timer ──
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $script:Config.RefreshIntervalSec * 1000

    # ── Assemble ──
    $form.Controls.Add($tabs); $form.Controls.Add($panelTop); $form.Controls.Add($statusBar)

    # ═══════════════════════════════════════════════════════════════════════
    # EVENT HANDLERS
    # ═══════════════════════════════════════════════════════════════════════

    $showSave = {
        param([string]$Name, [string]$Filter)
        $d = New-Object System.Windows.Forms.SaveFileDialog; $d.FileName = $Name; $d.Filter = $Filter
        if ($d.ShowDialog() -eq 'OK') { return $d.FileName }; return $null
    }

    # Helper: get current scope server (empty = all servers)
    $getScope = {
        $sel = $cmbScope.SelectedItem
        if (-not $sel -or $sel -eq '(All Servers)') { return '' }
        return $sel.ToString()
    }

    # Connect
    $btnConnect.Add_Click({
        $srv = $txtServer.Text.Trim()
        if (-not $srv) { [System.Windows.Forms.MessageBox]::Show('Enter server name.', 'Warning', 'OK', 'Warning'); return }
        $stLabel.Text = "Connecting to $srv..."
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            if ($script:ExSession) { Disconnect-ExchangeRemote -Session $script:ExSession }
            $script:ExSession = Connect-ExchangeRemote -Server $srv
            $lblStatus.Text = "Connected: $srv"; $lblStatus.ForeColor = [System.Drawing.Color]::Green
            $btnConnect.Text = 'Reconnect'; $stLabel.Text = "Connected to $srv"

            # Discover all Exchange transport servers and populate scope combo
            $stLabel.Text = 'Discovering transport servers...'
            try {
                $cmbScope.Items.Clear()
                $cmbScope.Items.Add('(All Servers)') | Out-Null
                $transportServers = Get-TransportService -ErrorAction Stop | Select-Object -ExpandProperty Name | Sort-Object
                $script:ExchangeServers = $transportServers
                foreach ($ts in $transportServers) { $cmbScope.Items.Add($ts) | Out-Null }
                $cmbScope.SelectedIndex = 0
                $stLabel.Text = "Connected to $srv — found $($transportServers.Count) transport server(s)"
            }
            catch {
                # Fallback: just add the connected server
                $cmbScope.Items.Add($srv) | Out-Null
                $stLabel.Text = "Connected to $srv (server discovery failed, added manually)"
            }
        }
        catch {
            $lblStatus.Text = 'Failed'; $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $stLabel.Text = "Failed: $_"
            [System.Windows.Forms.MessageBox]::Show("Connection failed:`n$_", 'Error', 'OK', 'Error')
        }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    # ── Dashboard ──
    $btnRefreshDash.Add_Click({
        $stLabel.Text = 'Loading dashboard...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $txtDash.Text = Get-DashboardData -Server (& $getScope)
            $stLabel.Text = 'Dashboard loaded'
        }
        catch { $stLabel.Text = "Dashboard error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    # ── Queues ──
    $loadQ = {
        $stLabel.Text = 'Loading queues...'; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $qs = Get-ExchangeQueues -Server (& $getScope) -Filter $txtQFilter.Text.Trim()
            $dgvQ.DataSource = [System.Collections.ArrayList]@($qs)
            $stLabel.Text = "$($qs.Count) queue(s)"
        }
        catch { $stLabel.Text = "Error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
    $loadM = {
        if ($dgvQ.SelectedRows.Count -eq 0) { return }
        $qid = $dgvQ.SelectedRows[0].Cells['Identity'].Value.ToString()
        $lblMsg.Text = "Messages: $qid"
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $ms = Get-QueueMessages -QueueIdentity $qid
            $dgvMsg.DataSource = [System.Collections.ArrayList]@($ms)
            $lblMsgCnt.Text = "$($ms.Count) msg(s)"
        } catch { $stLabel.Text = "Error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }

    $btnRefQ.Add_Click({ & $loadQ; & $loadErrors })
    $dgvQ.Add_SelectionChanged({ & $loadM })
    $timer.Add_Tick({ & $loadQ; & $loadErrors })
    $chkAutoRef.Add_CheckedChanged({ if ($chkAutoRef.Checked) { $timer.Start() } else { $timer.Stop() } })

    $btnRetryQ.Add_Click({
        if ($dgvQ.SelectedRows.Count -eq 0) { return }
        $qid = $dgvQ.SelectedRows[0].Cells['Identity'].Value.ToString()
        if ([System.Windows.Forms.MessageBox]::Show("Retry '$qid'?", 'Confirm', 'YesNo', 'Question') -eq 'Yes') {
            Invoke-QueueRetry -QueueIdentity $qid; & $loadQ
        }
    })

    $getSelMsg = {
        $ids = @(); foreach ($r in $dgvMsg.SelectedRows) { $ids += $r.Cells['Identity'].Value.ToString() }; return $ids
    }
    $btnRetryM.Add_Click({
        $s = & $getSelMsg; if ($s.Count -eq 0) { return }
        if ([System.Windows.Forms.MessageBox]::Show("Retry $($s.Count) msg(s)?", 'Confirm', 'YesNo', 'Question') -eq 'Yes') {
            $r = Invoke-MessageRetry -MessageIdentity $s
            $stLabel.Text = "Retry: $(($r|Where-Object Success).Count)/$($s.Count) ok"; & $loadM
        }
    })
    $btnSuspM.Add_Click({
        $s = & $getSelMsg; if ($s.Count -eq 0) { return }
        if ([System.Windows.Forms.MessageBox]::Show("Suspend $($s.Count) msg(s)?", 'Confirm', 'YesNo', 'Question') -eq 'Yes') {
            Suspend-QueueMessages -MessageIdentity $s; $stLabel.Text = "Suspended $($s.Count)"; & $loadM
        }
    })
    $btnRemM.Add_Click({
        $s = & $getSelMsg; if ($s.Count -eq 0) { return }
        $n = if ($chkNDR.Checked) {' with NDR'} else {''}
        if ([System.Windows.Forms.MessageBox]::Show("REMOVE $($s.Count) msg(s)$n?`nCannot undo!", 'Confirm', 'YesNo', 'Warning') -eq 'Yes') {
            Remove-QueueMessages -MessageIdentity $s -WithNDR:($chkNDR.Checked); $stLabel.Text = "Removed $($s.Count)"; & $loadM
        }
    })

    # ── Recent Errors (FAIL/DEFER/DSN) ──
    $script:LastQueueErrors = $null
    $loadErrors = {
        $stLabel.Text = 'Loading recent errors...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $srv = & $getScope
            $errEvents = @()
            foreach ($evType in @('FAIL','DEFER','DSN')) {
                $p = @{
                    Start = (Get-Date).AddDays(-1); End = Get-Date
                    EventId = $evType; ResultSize = 200; ErrorAction = 'Stop'
                }
                if ($srv) { $p['Server'] = $srv }
                try {
                    $errEvents += Get-MessageTrackingLog @p |
                        Select-Object Timestamp, EventId, Source, Sender,
                            @{N='Recipients';E={$_.Recipients -join '; '}},
                            MessageSubject, ServerHostname, ConnectorId,
                            RecipientStatus, MessageId, SourceContext
                } catch {}
            }
            $errEvents = $errEvents | Sort-Object Timestamp -Descending
            $script:LastQueueErrors = $errEvents

            $gridData = [System.Collections.ArrayList]::new()
            foreach ($e in $errEvents) {
                $gridData.Add([PSCustomObject]@{
                    Time        = $e.Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
                    Event       = $e.EventId
                    Sender      = $e.Sender
                    Recipients  = $e.Recipients
                    Subject     = $e.MessageSubject
                    Server      = $e.ServerHostname
                    Connector   = $e.ConnectorId
                    Error       = $e.RecipientStatus
                    MessageId   = $e.MessageId
                }) | Out-Null
            }
            $dgvErrors.DataSource = $gridData

            $failCnt = ($errEvents | Where-Object EventId -eq 'FAIL').Count
            $deferCnt = ($errEvents | Where-Object EventId -eq 'DEFER').Count
            $dsnCnt = ($errEvents | Where-Object EventId -eq 'DSN').Count
            $lblErrCnt.Text = "FAIL: $failCnt  |  DEFER: $deferCnt  |  DSN: $dsnCnt  |  Total: $($errEvents.Count)"
            $stLabel.Text = "Loaded $($errEvents.Count) error event(s)"
        }
        catch { $stLabel.Text = "Error loading errors: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }

    $btnRefErrors.Add_Click({ & $loadErrors })

    $btnExportErrors.Add_Click({
        if (-not $script:LastQueueErrors) { return }
        $f = & $showSave 'queue-errors.csv' 'CSV|*.csv|JSON|*.json'
        if ($f) {
            Export-ResultsToFile -Data $script:LastQueueErrors -FilePath $f -Format $(if ($f -match '\.json$'){'JSON'}else{'CSV'})
            $stLabel.Text = "Exported: $f"
        }
    })

    # ── Message Tracking ──
    $btnTrSearch.Add_Click({
        $stLabel.Text = 'Searching...'; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = @{ Start = $dtpTrStart.Value; End = $dtpTrEnd.Value }
            $srv = & $getScope; if ($srv) { $p['Server'] = $srv }
            $mid = $txtTrMsgId.Text.Trim(); if ($mid) { $p['MessageId'] = $mid }
            $snd = $txtTrSender.Text.Trim(); if ($snd) { $p['Sender'] = $snd }
            $rcp = $txtTrRecip.Text.Trim(); if ($rcp) { $p['Recipient'] = $rcp }
            $ev = $cmbTrEvent.SelectedItem.ToString(); if ($ev -ne '(All)') { $p['EventId'] = $ev }
            $sub = $txtTrSubject.Text.Trim(); if ($sub) { $p['Subject'] = $sub }

            $res = Trace-ExchangeMessage @p
            $script:LastTrackingResults = $res
            $dgvTrack.DataSource = [System.Collections.ArrayList]@($res)
            $lblTrCnt.Text = "$($res.Count) event(s)"
            $stLabel.Text = "Found $($res.Count) tracking event(s)"
        }
        catch { $stLabel.Text = "Error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnTrPath.Add_Click({
        if (-not $script:LastTrackingResults -or $dgvTrack.SelectedRows.Count -eq 0) { return }
        # Get MessageId of selected row, then show full path for that message
        $selMid = $dgvTrack.SelectedRows[0].Cells['MessageId'].Value
        if (-not $selMid) { [System.Windows.Forms.MessageBox]::Show('No MessageId in selected row.', 'Info', 'OK', 'Information'); return }

        $pathEvents = $script:LastTrackingResults | Where-Object { $_.MessageId -eq $selMid } | Sort-Object Timestamp
        $lines = @("MESSAGE PATH: $selMid", "Subject: $($pathEvents[0].MessageSubject)", "Sender: $($pathEvents[0].Sender)", "")
        $stepNum = 0
        foreach ($ev in $pathEvents) {
            $stepNum++
            $ts = $ev.Timestamp.ToString('HH:mm:ss.fff')
            $lines += "  [$stepNum] $ts  $($ev.EventId.PadRight(12)) $($ev.Source.PadRight(10)) $($ev.ServerHostname)"
            if ($ev.ConnectorId)    { $lines += "      Connector : $($ev.ConnectorId)" }
            if ($ev.Recipients)     { $lines += "      Recipients: $($ev.Recipients)" }
            if ($ev.RecipientStatus -and $ev.EventId -in @('FAIL','DSN','DEFER')) {
                $lines += "      Status    : $($ev.RecipientStatus)"
            }
            if ($ev.SourceContext)  { $lines += "      Context   : $($ev.SourceContext)" }
            if ($ev.ClientIp)       { $lines += "      ClientIP  : $($ev.ClientIp) ($($ev.ClientHostname))" }
        }

        $pathForm = New-Object System.Windows.Forms.Form
        $pathForm.Text = "Message Path — $selMid"
        $pathForm.Size = New-Object System.Drawing.Size(800, 600)
        $pathForm.StartPosition = 'CenterParent'
        $pathTxt = New-ConsoleTextBox; $pathTxt.Dock = 'Fill'
        $pathTxt.Text = ($lines -join "`r`n")
        $pathForm.Controls.Add($pathTxt)
        $pathForm.ShowDialog() | Out-Null
        $pathForm.Dispose()
    })

    $btnTrExport.Add_Click({
        if (-not $script:LastTrackingResults) { return }
        $f = & $showSave 'tracking.csv' 'CSV|*.csv|JSON|*.json'
        if ($f) { Export-ResultsToFile -Data $script:LastTrackingResults -FilePath $f -Format $(if ($f -match '\.json$'){'JSON'}else{'CSV'}); $stLabel.Text = "Exported: $f" }
    })

    # ── Protocol Logs ──
    $btnProtoBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog; $d.Description = 'Select Protocol Log directory'
        if ($d.ShowDialog() -eq 'OK') { $txtProtoPath.Text = $d.SelectedPath }
    })

    $btnProtoLoad.Add_Click({
        $path = $txtProtoPath.Text.Trim()
        if (-not $path) { [System.Windows.Forms.MessageBox]::Show('Enter log path.', 'Warning', 'OK', 'Warning'); return }
        $stLabel.Text = 'Parsing protocol logs...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $filter = $txtProtoFilter.Text.Trim()
            $max = [int]$nudProtoMax.Value
            $res = Parse-SmtpProtocolLog -LogPath $path -Filter $filter -MaxFiles $max
            $script:LastProtoResults = $res
            $dgvProto.DataSource = [System.Collections.ArrayList]@($res)
            $lblProtoCnt.Text = "$($res.Count) entries from $max file(s)"
            $stLabel.Text = "Parsed $($res.Count) protocol log entries"
        }
        catch { $stLabel.Text = "Error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnProtoExport.Add_Click({
        if (-not $script:LastProtoResults) { return }
        $f = & $showSave 'protocol-log.csv' 'CSV|*.csv|JSON|*.json'
        if ($f) { Export-ResultsToFile -Data $script:LastProtoResults -FilePath $f -Format $(if ($f -match '\.json$'){'JSON'}else{'CSV'}); $stLabel.Text = "Exported: $f" }
    })

    # ── Log Search ──
    $btnLogBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($d.ShowDialog() -eq 'OK') { $txtLogPath.Text = $d.SelectedPath }
    })

    $btnLogSearch.Add_Click({
        $lp = $txtLogPath.Text.Trim(); $pat = $txtLogPat.Text.Trim()
        if (-not $lp -or -not $pat) { [System.Windows.Forms.MessageBox]::Show('Enter path and pattern.', 'Warning', 'OK', 'Warning'); return }
        $stLabel.Text = "Searching $lp..."; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $res = Search-TransportLogs -LogPath $lp -Pattern $pat
            $script:LastLogResults = $res
            $gd = [System.Collections.ArrayList]::new()
            foreach ($r in $res) { $gd.Add([PSCustomObject]@{ File = $r.File; Line = $r.Line; Match = $r.Match; Date = $r.FileDate.ToString('yyyy-MM-dd HH:mm') }) | Out-Null }
            $dgvLog.DataSource = $gd
            $lblLogCnt.Text = "Results: $($res.Count) match(es)"
            $txtLogCtx.Text = ''; $stLabel.Text = "Found $($res.Count)"
        }
        catch { $stLabel.Text = "Error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $dgvLog.Add_SelectionChanged({
        if ($dgvLog.SelectedRows.Count -eq 0 -or -not $script:LastLogResults) { return }
        $i = $dgvLog.SelectedRows[0].Index
        if ($i -lt $script:LastLogResults.Count) { $txtLogCtx.Text = $script:LastLogResults[$i].Context }
    })

    $btnLogExport.Add_Click({
        if (-not $script:LastLogResults) { return }
        $f = & $showSave 'log-search.csv' 'CSV|*.csv|JSON|*.json'
        if ($f) { Export-ResultsToFile -Data $script:LastLogResults -FilePath $f -Format $(if ($f -match '\.json$'){'JSON'}else{'CSV'}); $stLabel.Text = "Exported: $f" }
    })

    # ── Header Analyzer ──
    $btnHdrFile.Add_Click({
        $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Email/Text|*.txt;*.eml;*.msg|All|*.*'
        if ($d.ShowDialog() -eq 'OK') { $txtHdrIn.Text = [System.IO.File]::ReadAllText($d.FileName) }
    })

    $btnHdrParse.Add_Click({
        $raw = $txtHdrIn.Text.Trim()
        if (-not $raw) { [System.Windows.Forms.MessageBox]::Show('Paste headers first.', 'Warning', 'OK', 'Warning'); return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = Parse-EmailHeaders -RawHeaders $raw
            $script:LastParsedHeaders = $p

            $info = @(
                "Message-ID  : $($p.MessageId)"
                "From        : $($p.From)"
                "To          : $($p.To)"
                "Subject     : $($p.Subject)"
                "Date        : $($p.Date)"
                "Return-Path : $($p.ReturnPath)"
                ""
                "SPF: $($p.SPF ?? 'N/A')  |  DKIM: $($p.DKIM ?? 'N/A')  |  DMARC: $($p.DMARC ?? 'N/A')"
                "Total hops: $($p.TotalHops)  |  Total delay: $([math]::Round($p.TotalDelayMs/1000, 2))s"
            )
            $txtHdrInfo.Text = $info -join "`r`n"

            $hd = [System.Collections.ArrayList]::new()
            $n = 0
            foreach ($hop in $p.Hops) {
                $n++
                $del = if ($hop.Delay) { "$([math]::Round($hop.Delay.TotalSeconds,2))s" } else { '-' }
                $ts = if ($hop.Timestamp -is [datetime]) { $hop.Timestamp.ToString('yyyy-MM-dd HH:mm:ss') } else { "$($hop.Timestamp)" }
                $hd.Add([PSCustomObject]@{
                    '#' = $n; From = $hop.From ?? '?'; By = $hop.By ?? '?'
                    Protocol = $hop.With ?? '?'; TLS = $hop.TLS; Timestamp = $ts; Delay = $del
                }) | Out-Null
            }
            $dgvHops.DataSource = $hd

            $xd = [System.Collections.ArrayList]::new()
            foreach ($k in ($p.XHeaders.Keys | Sort-Object)) {
                $xd.Add([PSCustomObject]@{ Header = $k; Value = $p.XHeaders[$k] }) | Out-Null
            }
            $dgvXH.DataSource = $xd

            $stLabel.Text = "Parsed: $($p.TotalHops) hops, $($p.XHeaders.Count) X-headers"
        }
        catch { $stLabel.Text = "Parse error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnHdrExport.Add_Click({
        if (-not $script:LastParsedHeaders) { return }
        $f = & $showSave 'header-analysis.csv' 'CSV|*.csv|JSON|*.json'
        if ($f) { Export-ResultsToFile -Data $script:LastParsedHeaders.Hops -FilePath $f -Format $(if ($f -match '\.json$'){'JSON'}else{'CSV'}); $stLabel.Text = "Exported: $f" }
    })

    # ── Reports ──
    $btnRunRpt.Add_Click({
        $rt = $cmbRpt.SelectedItem.ToString()
        $stLabel.Text = "Running $rt report..."; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $txtRpt.Text = Get-TransportReportData -ReportType $rt -Server (& $getScope)
            $stLabel.Text = "$rt report done"
        }
        catch { $stLabel.Text = "Report error: $_" }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })

    $btnSaveRpt.Add_Click({
        $c = $txtRpt.Text; if (-not $c) { return }
        $f = & $showSave 'transport-report.txt' 'Text|*.txt|JSON|*.json'
        if ($f) { $c | Set-Content -Path $f -Encoding UTF8; $stLabel.Text = "Saved: $f" }
    })

    # ── Form close ──
    $form.Add_FormClosing({
        $timer.Stop(); $timer.Dispose()
        if ($script:ExSession) { Disconnect-ExchangeRemote -Session $script:ExSession }
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
