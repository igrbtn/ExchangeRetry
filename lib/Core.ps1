#Requires -Version 5.1
<#
.SYNOPSIS
    Core Exchange transport functions library.
.DESCRIPTION
    Dot-sourced by ExchangeRetry.ps1 (GUI) and ExchangeTrace.ps1 (CLI).
    All functions return data only — no Write-Host.
#>

# ──────────────────────────────────────────────────────────────────────────────
# Connection
# ──────────────────────────────────────────────────────────────────────────────

function Connect-ExchangeRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Server,

        [PSCredential]$Credential,

        [string]$ConfigurationName = 'Microsoft.Exchange'
    )

    try {
        $sessionParams = @{
            ConfigurationName = $ConfigurationName
            ConnectionUri     = "http://$Server/PowerShell/"
            Authentication    = 'Kerberos'
            ErrorAction       = 'Stop'
        }
        if ($Credential) {
            $sessionParams['Credential'] = $Credential
        }

        $session = New-PSSession @sessionParams

        Import-PSSession -Session $session -DisableNameChecking -AllowClobber -ErrorAction Stop | Out-Null

        return $session
    }
    catch {
        throw "Failed to connect to Exchange server '$Server': $_"
    }
}

function Disconnect-ExchangeRemote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.Runspaces.PSSession]$Session
    )

    try {
        Remove-PSSession -Session $Session -ErrorAction Stop
    }
    catch {
        throw "Failed to disconnect Exchange session: $_"
    }
}

function Get-ExchangeTransportServers {
    [CmdletBinding()]
    param()

    try {
        $servers = Get-TransportService -ErrorAction Stop |
            Select-Object -ExpandProperty Name
        return $servers
    }
    catch {
        throw "Failed to retrieve transport servers: $_"
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Queue Operations
# ──────────────────────────────────────────────────────────────────────────────

function Get-ExchangeQueues {
    [CmdletBinding()]
    param(
        [string]$Server,
        [string]$Filter
    )

    try {
        $params = @{ ErrorAction = 'Stop' }
        if ($Server)  { $params['Server'] = $Server }
        if ($Filter)  { $params['Filter'] = $Filter }

        $queues = Get-Queue @params |
            Select-Object Identity, DeliveryType, Status, MessageCount,
                          NextHopDomain, LastError, NextRetryTime

        return $queues
    }
    catch {
        throw "Failed to get queues: $_"
    }
}

function Get-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$QueueIdentity,

        [int]$ResultSize = 500
    )

    try {
        $messages = Get-Message -Queue $QueueIdentity -ResultSize $ResultSize -ErrorAction Stop |
            Select-Object Identity, FromAddress, Status, Size, Subject,
                          DateReceived, LastError, SourceIP, SCL, RetryCount

        return $messages
    }
    catch {
        throw "Failed to get messages from queue '$QueueIdentity': $_"
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
        return [PSCustomObject]@{
            Queue   = $QueueIdentity
            Success = $true
            Error   = $null
        }
    }
    catch {
        return [PSCustomObject]@{
            Queue   = $QueueIdentity
            Success = $false
            Error   = $_.Exception.Message
        }
    }
}

function Invoke-MessageRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentities
    )

    $results = foreach ($id in $MessageIdentities) {
        try {
            Resume-Message -Identity $id -Confirm:$false -ErrorAction Stop
            [PSCustomObject]@{
                Identity = $id
                Success  = $true
                Error    = $null
            }
        }
        catch {
            [PSCustomObject]@{
                Identity = $id
                Success  = $false
                Error    = $_.Exception.Message
            }
        }
    }

    return $results
}

function Suspend-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentities
    )

    $results = foreach ($id in $MessageIdentities) {
        try {
            Suspend-Message -Identity $id -Confirm:$false -ErrorAction Stop
            [PSCustomObject]@{
                Identity = $id
                Success  = $true
                Error    = $null
            }
        }
        catch {
            [PSCustomObject]@{
                Identity = $id
                Success  = $false
                Error    = $_.Exception.Message
            }
        }
    }

    return $results
}

function Remove-QueueMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$MessageIdentities,

        [switch]$WithNDR
    )

    $results = foreach ($id in $MessageIdentities) {
        try {
            $removeParams = @{
                Identity    = $id
                WithNDR     = $WithNDR.IsPresent
                Confirm     = $false
                ErrorAction = 'Stop'
            }
            Remove-Message @removeParams
            [PSCustomObject]@{
                Identity = $id
                Success  = $true
                Error    = $null
            }
        }
        catch {
            [PSCustomObject]@{
                Identity = $id
                Success  = $false
                Error    = $_.Exception.Message
            }
        }
    }

    return $results
}

# ──────────────────────────────────────────────────────────────────────────────
# Message Tracking
# ──────────────────────────────────────────────────────────────────────────────

function Trace-ExchangeMessage {
    [CmdletBinding()]
    param(
        [string]$MessageId,
        [string]$Sender,
        [string]$Recipient,
        [string]$Server,
        [string]$EventId,
        [string]$ConnectorId,
        [string]$Source,
        [string]$Subject,
        [datetime]$Start,
        [datetime]$End,
        [int]$ResultSize = 1000
    )

    try {
        $params = @{
            ResultSize  = $ResultSize
            ErrorAction = 'Stop'
        }

        if ($MessageId)   { $params['MessageId']   = $MessageId }
        if ($Sender)      { $params['Sender']      = $Sender }
        if ($Recipient)   { $params['Recipients']   = $Recipient }
        if ($Server)      { $params['Server']      = $Server }
        if ($EventId)     { $params['EventId']     = $EventId }
        if ($ConnectorId) { $params['ConnectorId'] = $ConnectorId }
        if ($Source)       { $params['Source']      = $Source }
        if ($Start)        { $params['Start']      = $Start }
        if ($End)          { $params['End']        = $End }
        if ($Subject)      { $params['MessageSubject'] = $Subject }

        $logs = Get-MessageTrackingLog @params |
            Select-Object Timestamp, ClientIp, ClientHostname, OriginalClientIp,
                          ServerIp, ServerHostname, SourceContext, ConnectorId,
                          Source, EventId, InternalMessageId, MessageId,
                          @{N='Recipients';E={($_.Recipients -join '; ')}},
                          @{N='RecipientStatus';E={($_.RecipientStatus -join '; ')}},
                          Sender, ReturnPath, Directionality,
                          MessageSubject, TotalBytes, RecipientCount,
                          MessageLatency, MessageLatencyType

        return $logs
    }
    catch {
        throw "Message tracking query failed: $_"
    }
}

function Trace-CrossServerMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$MessageId,

        [Parameter(Mandatory)]
        [string[]]$Servers,

        [datetime]$Start,
        [datetime]$End,
        [int]$ResultSize = 1000
    )

    $allResults = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($srv in $Servers) {
        try {
            $params = @{
                MessageId   = $MessageId
                Server      = $srv
                ResultSize  = $ResultSize
                ErrorAction = 'Stop'
            }
            if ($Start) { $params['Start'] = $Start }
            if ($End)   { $params['End']   = $End }

            $logs = Get-MessageTrackingLog @params |
                Select-Object Timestamp, ClientIp, ClientHostname, OriginalClientIp,
                              ServerIp, ServerHostname, SourceContext, ConnectorId,
                              Source, EventId, InternalMessageId, MessageId,
                              @{N='Recipients';E={($_.Recipients -join '; ')}},
                              @{N='RecipientStatus';E={($_.RecipientStatus -join '; ')}},
                              Sender, ReturnPath, Directionality,
                              MessageSubject, TotalBytes, RecipientCount,
                              MessageLatency, MessageLatencyType

            foreach ($entry in $logs) {
                $allResults.Add($entry)
            }
        }
        catch {
            # Skip servers that error; include a diagnostic entry
            $allResults.Add([PSCustomObject]@{
                Timestamp       = [datetime]::MinValue
                ServerHostname  = $srv
                EventId         = 'QUERY_ERROR'
                Source           = 'CrossServerTrace'
                SourceContext    = $_.Exception.Message
                ClientIp        = $null; ClientHostname = $null; OriginalClientIp = $null
                ServerIp        = $null; ConnectorId = $null
                InternalMessageId = $null; MessageId = $MessageId
                Recipients      = $null; RecipientStatus = $null
                Sender          = $null; ReturnPath = $null; Directionality = $null
                MessageSubject  = $null; TotalBytes = $null; RecipientCount = $null
                MessageLatency  = $null; MessageLatencyType = $null
            })
        }
    }

    # Deduplicate by Timestamp + ServerHostname + EventId + Source
    $unique = $allResults |
        Sort-Object Timestamp, ServerHostname, EventId, Source -Unique

    return ($unique | Sort-Object Timestamp)
}

function Get-RecentErrors {
    [CmdletBinding()]
    param(
        [string]$Server,
        [int]$HoursBack = 24,
        [int]$ResultSize = 1000
    )

    $since = (Get-Date).AddHours(-$HoursBack)
    $allResults = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($eventType in @('FAIL', 'DEFER', 'DSN')) {
        try {
            $params = @{
                EventId     = $eventType
                Start       = $since
                ResultSize  = $ResultSize
                ErrorAction = 'Stop'
            }
            if ($Server) { $params['Server'] = $Server }

            $logs = Get-MessageTrackingLog @params |
                Select-Object Timestamp, ClientIp, ClientHostname, OriginalClientIp,
                              ServerHostname, ConnectorId, Source, EventId,
                              MessageId, Sender,
                              @{N='Recipients';E={($_.Recipients -join '; ')}},
                              @{N='RecipientStatus';E={($_.RecipientStatus -join '; ')}},
                              MessageSubject, SourceContext, TotalBytes

            foreach ($entry in $logs) {
                $allResults.Add($entry)
            }
        }
        catch {
            # Continue to next event type
        }
    }

    return ($allResults | Sort-Object Timestamp -Descending)
}

# ──────────────────────────────────────────────────────────────────────────────
# Header Parsing
# ──────────────────────────────────────────────────────────────────────────────

function Parse-EmailHeaders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RawHeaders
    )

    try {
        # Unfold continuation lines (lines starting with whitespace are continuations)
        $unfolded = $RawHeaders -replace "(\r?\n)[ \t]+", ' '
        $lines = $unfolded -split "\r?\n" | Where-Object { $_ -match '\S' }

        # Build header dictionary (multi-value aware)
        $headers = [ordered]@{}
        foreach ($line in $lines) {
            if ($line -match '^([A-Za-z0-9\-]+):\s*(.*)$') {
                $name  = $Matches[1]
                $value = $Matches[2].Trim()
                if ($headers.Contains($name)) {
                    if ($headers[$name] -is [System.Collections.Generic.List[string]]) {
                        $headers[$name].Add($value)
                    } else {
                        $prev = $headers[$name]
                        $headers[$name] = [System.Collections.Generic.List[string]]::new()
                        $headers[$name].Add($prev)
                        $headers[$name].Add($value)
                    }
                } else {
                    $headers[$name] = $value
                }
            }
        }

        # Helper to get single value
        $getFirst = {
            param($key)
            $v = $headers[$key]
            if ($null -eq $v) { return $null }
            if ($v -is [System.Collections.Generic.List[string]]) { return $v[0] }
            return $v
        }

        # Parse Received headers into hops
        $receivedRaw = @()
        if ($headers.Contains('Received')) {
            $rv = $headers['Received']
            if ($rv -is [System.Collections.Generic.List[string]]) {
                $receivedRaw = $rv
            } else {
                $receivedRaw = @($rv)
            }
        }

        $hops = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($rec in $receivedRaw) {
            $hop = [ordered]@{
                Raw       = $rec
                From      = $null
                By        = $null
                With      = $null
                For       = $null
                TLS       = $false
                Timestamp = $null
            }

            if ($rec -match 'from\s+(\S+)') {
                $hop['From'] = $Matches[1]
            }
            if ($rec -match 'by\s+(\S+)') {
                $hop['By'] = $Matches[1]
            }
            if ($rec -match 'with\s+(\S+)') {
                $hop['With'] = $Matches[1]
            }
            if ($rec -match 'for\s+(<[^>]+>|\S+)') {
                $hop['For'] = $Matches[1]
            }
            # TLS detection
            if ($rec -match '(?i)(TLS|STARTTLS|ESMTPS)') {
                $hop['TLS'] = $true
            }
            # Timestamp — typically after a semicolon
            if ($rec -match ';\s*(.+)$') {
                $tsString = $Matches[1].Trim()
                $parsed = $null
                if ([datetime]::TryParse($tsString, [ref]$parsed)) {
                    $hop['Timestamp'] = $parsed
                } else {
                    $hop['Timestamp'] = $tsString
                }
            }

            $hops.Add([PSCustomObject]$hop)
        }

        # Calculate delays between hops (Received headers are in reverse order)
        $reversedHops = [System.Collections.Generic.List[PSObject]]::new($hops)
        $reversedHops.Reverse()
        for ($i = 1; $i -lt $reversedHops.Count; $i++) {
            $prev = $reversedHops[$i - 1].Timestamp
            $curr = $reversedHops[$i].Timestamp
            if ($prev -is [datetime] -and $curr -is [datetime]) {
                $delay = $curr - $prev
                $reversedHops[$i] | Add-Member -NotePropertyName 'Delay' -NotePropertyValue $delay -Force
            } else {
                $reversedHops[$i] | Add-Member -NotePropertyName 'Delay' -NotePropertyValue $null -Force
            }
        }
        if ($reversedHops.Count -gt 0) {
            $reversedHops[0] | Add-Member -NotePropertyName 'Delay' -NotePropertyValue ([TimeSpan]::Zero) -Force
        }

        # Extract X-* headers
        $xHeaders = [ordered]@{}
        foreach ($key in $headers.Keys) {
            if ($key -match '^X-') {
                $xHeaders[$key] = $headers[$key]
            }
        }

        # Extract auth results
        $authResults = & $getFirst 'Authentication-Results'
        $spf = $null; $dkim = $null; $dmarc = $null

        if ($authResults) {
            if ($authResults -match '(?i)spf=(\S+)') { $spf = $Matches[1] }
            if ($authResults -match '(?i)dkim=(\S+)') { $dkim = $Matches[1] }
            if ($authResults -match '(?i)dmarc=(\S+)') { $dmarc = $Matches[1] }
        }
        # Also check dedicated headers
        $spfHeader = & $getFirst 'Received-SPF'
        if ($spfHeader -and -not $spf) {
            if ($spfHeader -match '^(\S+)') { $spf = $Matches[1] }
        }

        $result = [PSCustomObject]@{
            MessageId       = & $getFirst 'Message-ID'
            From            = & $getFirst 'From'
            To              = & $getFirst 'To'
            Subject         = & $getFirst 'Subject'
            Date            = & $getFirst 'Date'
            ReturnPath      = & $getFirst 'Return-Path'
            ContentType     = & $getFirst 'Content-Type'
            AuthResults     = $authResults
            SPF             = $spf
            DKIM            = $dkim
            DMARC           = $dmarc
            Hops            = $reversedHops
            XHeaders        = $xHeaders
            AllHeaders      = $headers
        }

        return $result
    }
    catch {
        throw "Failed to parse email headers: $_"
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Log Parsing
# ──────────────────────────────────────────────────────────────────────────────

function Parse-SmtpProtocolLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,

        [string]$Filter,

        [int]$MaxFiles = 50
    )

    try {
        if (-not (Test-Path $LogPath)) {
            throw "Log path not found: $LogPath"
        }

        $files = if ((Get-Item $LogPath).PSIsContainer) {
            Get-ChildItem -Path $LogPath -Filter '*.log' -File -ErrorAction Stop |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $MaxFiles
        } else {
            @(Get-Item $LogPath -ErrorAction Stop)
        }

        $results = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($file in $files) {
            $fieldNames = $null
            $content = Get-Content -Path $file.FullName -ErrorAction Stop

            foreach ($line in $content) {
                # Skip empty lines
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                # Parse #Fields header
                if ($line -match '^#Fields:\s*(.+)$') {
                    $fieldNames = $Matches[1] -split ','
                    $fieldNames = $fieldNames | ForEach-Object { $_.Trim() }
                    continue
                }

                # Skip other comment lines
                if ($line -match '^#') { continue }

                if ($null -eq $fieldNames) { continue }

                $values = $line -split ','
                $entry = [ordered]@{
                    _SourceFile = $file.Name
                }
                for ($i = 0; $i -lt $fieldNames.Count; $i++) {
                    $val = if ($i -lt $values.Count) { $values[$i].Trim() } else { $null }
                    $entry[$fieldNames[$i]] = $val
                }

                $obj = [PSCustomObject]$entry

                if ($Filter) {
                    $matchFound = $false
                    foreach ($prop in $obj.PSObject.Properties) {
                        if ($prop.Value -and $prop.Value.ToString() -like "*$Filter*") {
                            $matchFound = $true
                            break
                        }
                    }
                    if (-not $matchFound) { continue }
                }

                $results.Add($obj)
            }
        }

        return $results
    }
    catch {
        throw "Failed to parse SMTP protocol logs: $_"
    }
}

function Search-TransportLogs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,

        [Parameter(Mandatory)]
        [string]$Pattern,

        [int]$ContextLines = 2,

        [int]$MaxFiles = 50
    )

    try {
        if (-not (Test-Path $LogPath)) {
            throw "Log path not found: $LogPath"
        }

        $files = if ((Get-Item $LogPath).PSIsContainer) {
            Get-ChildItem -Path $LogPath -Filter '*.log' -File -ErrorAction Stop |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $MaxFiles
        } else {
            @(Get-Item $LogPath -ErrorAction Stop)
        }

        $results = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($file in $files) {
            $lines = Get-Content -Path $file.FullName -ErrorAction Stop
            for ($i = 0; $i -lt $lines.Count; $i++) {
                if ($lines[$i] -match $Pattern) {
                    $ctxStart = [Math]::Max(0, $i - $ContextLines)
                    $ctxEnd   = [Math]::Min($lines.Count - 1, $i + $ContextLines)

                    $contextBlock = for ($j = $ctxStart; $j -le $ctxEnd; $j++) {
                        $prefix = if ($j -eq $i) { '>>>' } else { '   ' }
                        "$prefix $($j + 1): $($lines[$j])"
                    }

                    $results.Add([PSCustomObject]@{
                        File       = $file.Name
                        FilePath   = $file.FullName
                        LineNumber = $i + 1
                        MatchLine  = $lines[$i]
                        Context    = ($contextBlock -join "`n")
                    })
                }
            }
        }

        return $results
    }
    catch {
        throw "Failed to search transport logs: $_"
    }
}

function Parse-ConnectivityLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$LogPath,

        [string]$Filter,

        [int]$MaxFiles = 50
    )

    try {
        if (-not (Test-Path $LogPath)) {
            throw "Log path not found: $LogPath"
        }

        $files = if ((Get-Item $LogPath).PSIsContainer) {
            Get-ChildItem -Path $LogPath -Filter '*.log' -File -ErrorAction Stop |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $MaxFiles
        } else {
            @(Get-Item $LogPath -ErrorAction Stop)
        }

        $results = [System.Collections.Generic.List[PSObject]]::new()

        foreach ($file in $files) {
            $fieldNames = $null
            $content = Get-Content -Path $file.FullName -ErrorAction Stop

            foreach ($line in $content) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                if ($line -match '^#Fields:\s*(.+)$') {
                    $fieldNames = $Matches[1] -split ','
                    $fieldNames = $fieldNames | ForEach-Object { $_.Trim() }
                    continue
                }

                if ($line -match '^#') { continue }
                if ($null -eq $fieldNames) { continue }

                $values = $line -split ','
                $entry = [ordered]@{
                    _SourceFile = $file.Name
                }
                for ($i = 0; $i -lt $fieldNames.Count; $i++) {
                    $val = if ($i -lt $values.Count) { $values[$i].Trim() } else { $null }
                    $entry[$fieldNames[$i]] = $val
                }

                $obj = [PSCustomObject]$entry

                if ($Filter) {
                    $matchFound = $false
                    foreach ($prop in $obj.PSObject.Properties) {
                        if ($prop.Value -and $prop.Value.ToString() -like "*$Filter*") {
                            $matchFound = $true
                            break
                        }
                    }
                    if (-not $matchFound) { continue }
                }

                $results.Add($obj)
            }
        }

        return $results
    }
    catch {
        throw "Failed to parse connectivity logs: $_"
    }
}

# ──────────────────────────────────────────────────────────────────────────────
# Reports
# ──────────────────────────────────────────────────────────────────────────────

function Get-TransportReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Full','Queues','Connectors','AgentLog','RoutingTable','DSN','Summary','Pipeline','BackPressure')]
        [string]$ReportType,

        [string]$Server
    )

    $sb = [System.Text.StringBuilder]::new()

    # ── Helper: section separator ──
    $sep = { param($title)
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(('=' * 80))
        [void]$sb.AppendLine("  $title")
        [void]$sb.AppendLine(('=' * 80))
    }

    # ── Queues ──
    $buildQueues = {
        & $sep 'QUEUE STATUS'
        try {
            $qParams = @{ ErrorAction = 'Stop' }
            if ($Server) { $qParams['Server'] = $Server }
            $queues = Get-Queue @qParams

            # Summary by status
            $byStatus = $queues | Group-Object Status
            [void]$sb.AppendLine('Status Summary:')
            foreach ($g in $byStatus) {
                $totalMsg = ($g.Group | Measure-Object -Property MessageCount -Sum).Sum
                [void]$sb.AppendLine("  $($g.Name): $($g.Count) queue(s), $totalMsg message(s)")
            }

            # Retry queues with errors
            $retryQueues = $queues | Where-Object { $_.Status -eq 'Retry' -or $_.MessageCount -gt 0 }
            if ($retryQueues) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('Retry / Non-Empty Queues:')
                foreach ($q in $retryQueues) {
                    [void]$sb.AppendLine("  Queue: $($q.Identity)")
                    [void]$sb.AppendLine("    DeliveryType : $($q.DeliveryType)")
                    [void]$sb.AppendLine("    Status       : $($q.Status)")
                    [void]$sb.AppendLine("    Messages     : $($q.MessageCount)")
                    [void]$sb.AppendLine("    NextHop      : $($q.NextHopDomain)")
                    [void]$sb.AppendLine("    NextRetry    : $($q.NextRetryTime)")
                    if ($q.LastError) {
                        [void]$sb.AppendLine("    LastError    : $($q.LastError)")
                    }
                }
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── Connectors ──
    $buildConnectors = {
        & $sep 'SEND CONNECTORS'
        try {
            $sendConn = Get-SendConnector -ErrorAction Stop
            foreach ($c in $sendConn) {
                [void]$sb.AppendLine("  [$($c.Name)]")
                [void]$sb.AppendLine("    Enabled        : $($c.Enabled)")
                [void]$sb.AppendLine("    AddressSpaces  : $(($c.AddressSpaces -join ', '))")
                [void]$sb.AppendLine("    SmartHosts     : $(($c.SmartHosts -join ', '))")
                [void]$sb.AppendLine("    TlsDomain      : $($c.TlsDomain)")
                [void]$sb.AppendLine("    RequireTLS     : $($c.RequireTLS)")
                [void]$sb.AppendLine("    MaxMessageSize : $($c.MaxMessageSize)")
                [void]$sb.AppendLine('')
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }

        & $sep 'RECEIVE CONNECTORS'
        try {
            $rcvParams = @{ ErrorAction = 'Stop' }
            if ($Server) { $rcvParams['Server'] = $Server }
            $rcvConn = Get-ReceiveConnector @rcvParams
            foreach ($c in $rcvConn) {
                [void]$sb.AppendLine("  [$($c.Name)]")
                [void]$sb.AppendLine("    Enabled           : $($c.Enabled)")
                [void]$sb.AppendLine("    Bindings          : $(($c.Bindings -join ', '))")
                [void]$sb.AppendLine("    RemoteIPRanges    : $(($c.RemoteIPRanges -join ', '))")
                [void]$sb.AppendLine("    AuthMechanism     : $(($c.AuthMechanism -join ', '))")
                [void]$sb.AppendLine("    PermissionGroups  : $(($c.PermissionGroups -join ', '))")
                [void]$sb.AppendLine("    MaxMessageSize    : $($c.MaxMessageSize)")
                [void]$sb.AppendLine("    ProtocolLogging   : $($c.ProtocolLoggingLevel)")
                [void]$sb.AppendLine('')
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── AgentLog ──
    $buildAgentLog = {
        & $sep 'TRANSPORT AGENTS'
        try {
            $agents = Get-TransportAgent -ErrorAction Stop | Sort-Object Priority
            [void]$sb.AppendLine('Registered Agents (by priority):')
            foreach ($a in $agents) {
                [void]$sb.AppendLine("  [$($a.Priority)] $($a.Identity) (Enabled: $($a.Enabled))")
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR getting agents: $_")
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Recent Agent Log Entries (last 100):')
        try {
            $agentParams = @{ ErrorAction = 'Stop' }
            if ($Server) { $agentParams['Server'] = $Server }
            $agentLogs = Get-AgentLog @agentParams | Select-Object -Last 100
            foreach ($entry in $agentLogs) {
                [void]$sb.AppendLine("  $($entry.Timestamp) | $($entry.Agent) | $($entry.Event) | $($entry.Action) | $($entry.SmtpResponse)")
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR getting agent logs: $_")
        }
    }

    # ── RoutingTable / TransportConfig ──
    $buildRoutingTable = {
        & $sep 'TRANSPORT CONFIGURATION'
        try {
            $tc = Get-TransportConfig -ErrorAction Stop
            [void]$sb.AppendLine('Global Transport Config:')
            [void]$sb.AppendLine("  MaxReceiveSize         : $($tc.MaxReceiveSize)")
            [void]$sb.AppendLine("  MaxSendSize            : $($tc.MaxSendSize)")
            [void]$sb.AppendLine("  MaxRecipientEnvelopeLimit : $($tc.MaxRecipientEnvelopeLimit)")
            [void]$sb.AppendLine("  ShadowRedundancyEnabled : $($tc.ShadowRedundancyEnabled)")
            [void]$sb.AppendLine("  SafetyNetHoldTime      : $($tc.SafetyNetHoldTime)")
            [void]$sb.AppendLine("  TLSSendDomainSecureList: $(($tc.TLSSendDomainSecureList -join ', '))")
            [void]$sb.AppendLine("  TLSReceiveDomainSecureList: $(($tc.TLSReceiveDomainSecureList -join ', '))")
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }

        & $sep 'TRANSPORT SERVICE SETTINGS'
        try {
            $tsParams = @{ ErrorAction = 'Stop' }
            if ($Server) { $tsParams['Identity'] = $Server }
            $ts = Get-TransportService @tsParams
            foreach ($s in $ts) {
                [void]$sb.AppendLine("  Server: $($s.Name)")
                [void]$sb.AppendLine("    MessageTrackingLogPath        : $($s.MessageTrackingLogPath)")
                [void]$sb.AppendLine("    SendProtocolLogPath           : $($s.SendProtocolLogPath)")
                [void]$sb.AppendLine("    ReceiveProtocolLogPath        : $($s.ReceiveProtocolLogPath)")
                [void]$sb.AppendLine("    ConnectivityLogPath           : $($s.ConnectivityLogPath)")
                [void]$sb.AppendLine("    PipelineTracingPath           : $($s.PipelineTracingPath)")
                [void]$sb.AppendLine("    RoutingTableLogPath           : $($s.RoutingTableLogPath)")
                [void]$sb.AppendLine("    MessageRetryInterval          : $($s.MessageRetryInterval)")
                [void]$sb.AppendLine("    MailboxDeliveryQueueRetryInterval : $($s.MailboxDeliveryQueueRetryInterval)")
                [void]$sb.AppendLine("    OutboundConnectionFailureRetryInterval : $($s.OutboundConnectionFailureRetryInterval)")
                [void]$sb.AppendLine("    TransientFailureRetryInterval : $($s.TransientFailureRetryInterval)")
                [void]$sb.AppendLine("    TransientFailureRetryCount    : $($s.TransientFailureRetryCount)")
                [void]$sb.AppendLine("    MessageExpirationTimeout      : $($s.MessageExpirationTimeout)")
                [void]$sb.AppendLine('')
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── DSN ──
    $buildDSN = {
        & $sep 'DSN CONFIGURATION'
        try {
            $tc = Get-TransportConfig -ErrorAction Stop
            [void]$sb.AppendLine("  GenerateCopyOfDSNFor           : $(($tc.GenerateCopyOfDSNFor -join ', '))")
            [void]$sb.AppendLine("  DSNConversionMode              : $($tc.DSNConversionMode)")
            [void]$sb.AppendLine("  ExternalDsnLanguageDetectionEnabled : $($tc.ExternalDsnLanguageDetectionEnabled)")
            [void]$sb.AppendLine("  InternalDsnLanguageDetectionEnabled : $($tc.InternalDsnLanguageDetectionEnabled)")
            [void]$sb.AppendLine("  ExternalDsnDefaultLanguage     : $($tc.ExternalDsnDefaultLanguage)")
            [void]$sb.AppendLine("  ExternalDsnMaxMessageAttachSize: $($tc.ExternalDsnMaxMessageAttachSize)")
            [void]$sb.AppendLine("  InternalDsnMaxMessageAttachSize: $($tc.InternalDsnMaxMessageAttachSize)")
            [void]$sb.AppendLine("  ExternalDsnReportingAuthority  : $($tc.ExternalDsnReportingAuthority)")
            [void]$sb.AppendLine("  ExternalDsnSendHtml            : $($tc.ExternalDsnSendHtml)")
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Recent DSN Events (last 24h):')
        try {
            $dsnParams = @{
                EventId     = 'DSN'
                Start       = (Get-Date).AddHours(-24)
                ResultSize  = 200
                ErrorAction = 'Stop'
            }
            if ($Server) { $dsnParams['Server'] = $Server }
            $dsnEvents = Get-MessageTrackingLog @dsnParams
            foreach ($d in $dsnEvents) {
                [void]$sb.AppendLine("  $($d.Timestamp) | $($d.Sender) -> $(($d.Recipients -join ';')) | $($d.MessageSubject) | $($d.SourceContext)")
            }
            if (-not $dsnEvents) {
                [void]$sb.AppendLine('  (none)')
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── Summary ──
    $buildSummary = {
        & $sep 'DELIVERY SUMMARY (last 24 hours)'
        try {
            $since = (Get-Date).AddHours(-24)
            $baseParams = @{
                Start       = $since
                ResultSize  = 'Unlimited'
                ErrorAction = 'Stop'
            }
            if ($Server) { $baseParams['Server'] = $Server }

            # Gather all tracking entries
            $allLogs = Get-MessageTrackingLog @baseParams

            $eventCounts = $allLogs | Group-Object EventId | Sort-Object Count -Descending
            [void]$sb.AppendLine('Event Counts:')
            foreach ($ec in $eventCounts) {
                [void]$sb.AppendLine("  $($ec.Name): $($ec.Count)")
            }

            $delivered = ($eventCounts | Where-Object { $_.Name -eq 'DELIVER' }).Count
            $sent      = ($eventCounts | Where-Object { $_.Name -eq 'SEND' }).Count
            $received  = ($eventCounts | Where-Object { $_.Name -eq 'RECEIVE' }).Count
            $failed    = ($eventCounts | Where-Object { $_.Name -eq 'FAIL' }).Count
            $total = $delivered + $sent + $failed
            $deliveryRate = if ($total -gt 0) { [math]::Round(($delivered + $sent) / $total * 100, 1) } else { 'N/A' }

            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("Delivery Rate: $deliveryRate%")

            # By server breakdown
            $byServer = $allLogs | Group-Object ServerHostname
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('By Server:')
            foreach ($s in $byServer) {
                [void]$sb.AppendLine("  $($s.Name): $($s.Count) events")
            }

            # Top 20 senders
            $topSenders = $allLogs |
                Where-Object { $_.Sender } |
                Group-Object Sender |
                Sort-Object Count -Descending |
                Select-Object -First 20

            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Top 20 Senders:')
            foreach ($ts in $topSenders) {
                [void]$sb.AppendLine("  $($ts.Count) msg  $($ts.Name)")
            }

            # Recent failures
            $recentFails = $allLogs |
                Where-Object { $_.EventId -in @('FAIL','DEFER') } |
                Sort-Object Timestamp -Descending |
                Select-Object -First 20

            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Recent Failures/Defers (last 20):')
            foreach ($f in $recentFails) {
                [void]$sb.AppendLine("  $($f.Timestamp) | $($f.EventId) | $($f.Sender) -> $(($f.Recipients -join ';')) | $($f.SourceContext)")
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── Pipeline ──
    $buildPipeline = {
        & $sep 'TRANSPORT PIPELINE (Agents by Priority)'
        try {
            $agents = Get-TransportAgent -ErrorAction Stop | Sort-Object Priority
            foreach ($a in $agents) {
                [void]$sb.AppendLine("  [$($a.Priority)] $($a.Identity)")
                [void]$sb.AppendLine("    Enabled      : $($a.Enabled)")
                [void]$sb.AppendLine("    AssemblyPath : $($a.AssemblyPath)")
                [void]$sb.AppendLine('')
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── BackPressure ──
    $buildBackPressure = {
        & $sep 'BACK PRESSURE / RESOURCE THROTTLING'
        try {
            $srvName = if ($Server) { $Server } else { $env:COMPUTERNAME }
            $diag = Get-ExchangeDiagnosticInfo -Server $srvName -Process EdgeTransport -Component ResourceThrottling -ErrorAction Stop
            [void]$sb.AppendLine('Resource Throttling Diagnostics:')
            [void]$sb.AppendLine($diag)
        }
        catch {
            [void]$sb.AppendLine("  ERROR getting diagnostics: $_")
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Queue Database Paths:')
        try {
            $tsParams = @{ ErrorAction = 'Stop' }
            if ($Server) { $tsParams['Identity'] = $Server }
            $ts = Get-TransportService @tsParams
            foreach ($s in $ts) {
                [void]$sb.AppendLine("  Server: $($s.Name)")
                [void]$sb.AppendLine("    QueueDatabasePath    : $($s.QueueDatabasePath)")
                [void]$sb.AppendLine("    QueueDatabaseLoggingPath : $($s.QueueDatabaseLoggingPath)")
            }
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    # ── Build the requested report ──
    [void]$sb.AppendLine("Exchange Transport Report: $ReportType")
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    if ($Server) { [void]$sb.AppendLine("Server: $Server") }

    switch ($ReportType) {
        'Queues'       { & $buildQueues }
        'Connectors'   { & $buildConnectors }
        'AgentLog'     { & $buildAgentLog }
        'RoutingTable' { & $buildRoutingTable }
        'DSN'          { & $buildDSN }
        'Summary'      { & $buildSummary }
        'Pipeline'     { & $buildPipeline }
        'BackPressure' { & $buildBackPressure }
        'Full' {
            & $buildQueues
            & $buildConnectors
            & $buildAgentLog
            & $buildRoutingTable
            & $buildDSN
            & $buildSummary
            & $buildPipeline
            & $buildBackPressure
        }
    }

    return $sb.ToString()
}

function Get-DashboardData {
    [CmdletBinding()]
    param(
        [string]$Server
    )

    $sb = [System.Text.StringBuilder]::new()

    [void]$sb.AppendLine('EXCHANGE TRANSPORT DASHBOARD')
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine(('-' * 60))

    # ── Queue Summary ──
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('QUEUE OVERVIEW')
    try {
        $qParams = @{ ErrorAction = 'Stop' }
        if ($Server) { $qParams['Server'] = $Server }
        $queues = Get-Queue @qParams

        $totalMessages = ($queues | Measure-Object -Property MessageCount -Sum).Sum
        [void]$sb.AppendLine("  Total Queues   : $($queues.Count)")
        [void]$sb.AppendLine("  Total Messages : $totalMessages")

        $byStatus = $queues | Group-Object Status
        foreach ($g in $byStatus) {
            $msgs = ($g.Group | Measure-Object -Property MessageCount -Sum).Sum
            [void]$sb.AppendLine("  $($g.Name): $($g.Count) queue(s), $msgs msg(s)")
        }

        # Retry details
        $retryQ = $queues | Where-Object { $_.Status -eq 'Retry' }
        if ($retryQ) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('  Retry Queue Details:')
            foreach ($r in $retryQ) {
                [void]$sb.AppendLine("    $($r.Identity) | $($r.MessageCount) msg | NextRetry: $($r.NextRetryTime) | $($r.LastError)")
            }
        }
    }
    catch {
        [void]$sb.AppendLine("  ERROR: $_")
    }

    # ── Delivery Stats ──
    $buildStats = { param($label, $hours)
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("DELIVERY STATS (last $label)")
        try {
            $since = (Get-Date).AddHours(-$hours)
            $baseParams = @{
                Start       = $since
                ResultSize  = 'Unlimited'
                ErrorAction = 'Stop'
            }
            if ($Server) { $baseParams['Server'] = $Server }
            $logs = Get-MessageTrackingLog @baseParams

            $counts = @{}
            foreach ($entry in $logs) {
                $eid = $entry.EventId
                if ($counts.ContainsKey($eid)) { $counts[$eid]++ } else { $counts[$eid] = 1 }
            }

            $received  = if ($counts['RECEIVE'])  { $counts['RECEIVE'] }  else { 0 }
            $delivered = if ($counts['DELIVER'])   { $counts['DELIVER'] }  else { 0 }
            $sent      = if ($counts['SEND'])      { $counts['SEND'] }     else { 0 }
            $failed    = if ($counts['FAIL'])       { $counts['FAIL'] }    else { 0 }
            $deferred  = if ($counts['DEFER'])      { $counts['DEFER'] }   else { 0 }
            $dsn       = if ($counts['DSN'])        { $counts['DSN'] }     else { 0 }

            $total = $delivered + $sent + $failed
            $rate = if ($total -gt 0) { [math]::Round(($delivered + $sent) / $total * 100, 1) } else { 'N/A' }

            [void]$sb.AppendLine("  Received  : $received")
            [void]$sb.AppendLine("  Delivered : $delivered")
            [void]$sb.AppendLine("  Sent      : $sent")
            [void]$sb.AppendLine("  Failed    : $failed")
            [void]$sb.AppendLine("  Deferred  : $deferred")
            [void]$sb.AppendLine("  DSN       : $dsn")
            [void]$sb.AppendLine("  Delivery Rate : $rate%")
        }
        catch {
            [void]$sb.AppendLine("  ERROR: $_")
        }
    }

    & $buildStats '1 hour' 1
    & $buildStats '24 hours' 24

    # ── Connector Health ──
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('CONNECTOR HEALTH')
    try {
        $sendConn = Get-SendConnector -ErrorAction Stop
        [void]$sb.AppendLine("  Send Connectors: $($sendConn.Count)")
        $enabledSend = ($sendConn | Where-Object { $_.Enabled }).Count
        [void]$sb.AppendLine("    Enabled  : $enabledSend")
        [void]$sb.AppendLine("    Disabled : $($sendConn.Count - $enabledSend)")
    }
    catch {
        [void]$sb.AppendLine("  Send Connectors: ERROR - $_")
    }

    try {
        $rcvParams = @{ ErrorAction = 'Stop' }
        if ($Server) { $rcvParams['Server'] = $Server }
        $rcvConn = Get-ReceiveConnector @rcvParams
        [void]$sb.AppendLine("  Receive Connectors: $($rcvConn.Count)")
        $enabledRcv = ($rcvConn | Where-Object { $_.Enabled }).Count
        [void]$sb.AppendLine("    Enabled  : $enabledRcv")
        [void]$sb.AppendLine("    Disabled : $($rcvConn.Count - $enabledRcv)")
    }
    catch {
        [void]$sb.AppendLine("  Receive Connectors: ERROR - $_")
    }

    return $sb.ToString()
}

# ──────────────────────────────────────────────────────────────────────────────
# Export
# ──────────────────────────────────────────────────────────────────────────────

function Export-ResultsToFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Data,

        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter(Mandatory)]
        [ValidateSet('CSV','JSON')]
        [string]$Format
    )

    try {
        $dir = Split-Path -Path $FilePath -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }

        switch ($Format) {
            'CSV' {
                $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            }
            'JSON' {
                $Data | ConvertTo-Json -Depth 10 -ErrorAction Stop |
                    Set-Content -Path $FilePath -Encoding UTF8 -ErrorAction Stop
            }
        }

        return [PSCustomObject]@{
            FilePath = (Resolve-Path $FilePath).Path
            Format   = $Format
            Records  = @($Data).Count
            Size     = (Get-Item $FilePath).Length
            Success  = $true
        }
    }
    catch {
        throw "Failed to export results to '$FilePath' as $Format: $_"
    }
}
