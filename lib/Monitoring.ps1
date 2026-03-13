# Monitoring.ps1 - Monitoring, settings, caching, and logging functions for ExchangeRetry
# Dot-source this file from the main script.

$script:AppDataPath = Join-Path $env:APPDATA 'ExchangeRetry'

$script:Cache = @{}

#region Settings Persistence

function Initialize-AppData {
    [CmdletBinding()]
    param()

    try {
        if (-not (Test-Path $script:AppDataPath)) {
            New-Item -Path $script:AppDataPath -ItemType Directory -Force | Out-Null
        }

        $settingsFile = Join-Path $script:AppDataPath 'settings.json'
        if (-not (Test-Path $settingsFile)) {
            $defaults = @{
                LastServer         = ''
                LastScope           = 'All'
                WindowWidth         = 1200
                WindowHeight        = 800
                SplitterPositions   = @{
                    Main   = 300
                    Detail = 400
                }
                RefreshIntervalSec  = 30
                AlertThresholds     = @{
                    MaxQueueDepth   = 100
                    MinDeliveryRate = 95
                    MaxRetryQueues  = 0
                }
                AlertSoundEnabled   = $true
                Theme               = 'Light'
                RecentServers       = @()
            }
            $defaults | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsFile -Encoding UTF8
        }

        $logFile = Join-Path $script:AppDataPath 'operator-log.csv'
        if (-not (Test-Path $logFile)) {
            [PSCustomObject]@{
                Timestamp = $null
                User      = $null
                Action    = $null
                Target    = $null
                Details   = $null
            } | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
            # Remove the dummy data row, keep header only
            $header = Get-Content -Path $logFile -TotalCount 1
            Set-Content -Path $logFile -Value $header -Encoding UTF8
        }
    }
    catch {
        Write-Verbose "Initialize-AppData error: $_"
    }
}

function Get-AppSettings {
    [CmdletBinding()]
    param()

    $defaults = @{
        LastServer         = ''
        LastScope           = 'All'
        WindowWidth         = 1200
        WindowHeight        = 800
        SplitterPositions   = @{
            Main   = 300
            Detail = 400
        }
        RefreshIntervalSec  = 30
        AlertThresholds     = @{
            MaxQueueDepth   = 100
            MinDeliveryRate = 95
            MaxRetryQueues  = 0
        }
        AlertSoundEnabled   = $true
        Theme               = 'Light'
        RecentServers       = @()
    }

    try {
        $settingsFile = Join-Path $script:AppDataPath 'settings.json'
        if (-not (Test-Path $settingsFile)) {
            return $defaults
        }

        $json = Get-Content -Path $settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json

        $settings = @{}
        foreach ($key in $defaults.Keys) {
            $val = $json.PSObject.Properties[$key]
            if ($null -ne $val) {
                $raw = $val.Value
                # Convert PSCustomObject to hashtable for nested objects
                if ($raw -is [System.Management.Automation.PSCustomObject]) {
                    $ht = @{}
                    foreach ($p in $raw.PSObject.Properties) {
                        $ht[$p.Name] = $p.Value
                    }
                    $settings[$key] = $ht
                }
                elseif ($raw -is [System.Object[]]) {
                    $settings[$key] = @($raw)
                }
                else {
                    $settings[$key] = $raw
                }
            }
            else {
                $settings[$key] = $defaults[$key]
            }
        }

        return $settings
    }
    catch {
        Write-Verbose "Get-AppSettings error: $_"
        return $defaults
    }
}

function Save-AppSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Settings
    )

    try {
        Initialize-AppData
        $settingsFile = Join-Path $script:AppDataPath 'settings.json'
        $Settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsFile -Encoding UTF8
    }
    catch {
        Write-Verbose "Save-AppSettings error: $_"
    }
}

function Update-RecentServers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Server
    )

    try {
        $settings = Get-AppSettings

        [System.Collections.ArrayList]$list = @()
        if ($settings.RecentServers) {
            foreach ($s in $settings.RecentServers) {
                $list.Add($s) | Out-Null
            }
        }

        # Remove duplicate if exists
        $existing = $list | Where-Object { $_ -eq $Server }
        if ($existing) {
            $list.Remove($Server)
        }

        # Insert at front
        $list.Insert(0, $Server)

        # Cap at 10
        if ($list.Count -gt 10) {
            $list = [System.Collections.ArrayList]@($list[0..9])
        }

        $settings.RecentServers = @($list)
        Save-AppSettings -Settings $settings
    }
    catch {
        Write-Verbose "Update-RecentServers error: $_"
    }
}

#endregion

#region Cache

function Get-CachedData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [int]$MaxAgeSec = 300
    )

    try {
        if ($script:Cache.ContainsKey($Key)) {
            $entry = $script:Cache[$Key]
            $age = (Get-Date) - $entry.Timestamp
            if ($age.TotalSeconds -le $MaxAgeSec) {
                return $entry.Value
            }
        }
    }
    catch {
        Write-Verbose "Get-CachedData error: $_"
    }

    return $null
}

function Set-CachedData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Key,

        [Parameter(Mandatory)]
        $Value
    )

    try {
        $script:Cache[$Key] = @{
            Value     = $Value
            Timestamp = Get-Date
        }
    }
    catch {
        Write-Verbose "Set-CachedData error: $_"
    }
}

function Clear-Cache {
    [CmdletBinding()]
    param()

    $script:Cache = @{}
}

#endregion

#region Operator Action Log

function Write-OperatorLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Action,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Target,

        [string]$Details = '',

        [string]$User = $env:USERNAME
    )

    try {
        Initialize-AppData
        $logFile = Join-Path $script:AppDataPath 'operator-log.csv'

        [PSCustomObject]@{
            Timestamp = (Get-Date -Format 'o')
            User      = $User
            Action    = $Action
            Target    = $Target
            Details   = $Details
        } | Export-Csv -Path $logFile -Append -NoTypeInformation -Encoding UTF8
    }
    catch {
        Write-Verbose "Write-OperatorLog error: $_"
    }
}

function Get-OperatorLog {
    [CmdletBinding()]
    param(
        [int]$Last = 100
    )

    try {
        $logFile = Join-Path $script:AppDataPath 'operator-log.csv'
        if (-not (Test-Path $logFile)) {
            return @()
        }

        $entries = Import-Csv -Path $logFile -Encoding UTF8
        if (-not $entries) {
            return @()
        }

        $entries = @($entries)
        if ($entries.Count -le $Last) {
            return $entries
        }

        return $entries[($entries.Count - $Last)..($entries.Count - 1)]
    }
    catch {
        Write-Verbose "Get-OperatorLog error: $_"
        return @()
    }
}

function Clear-OperatorLog {
    [CmdletBinding()]
    param()

    try {
        Initialize-AppData
        $logFile = Join-Path $script:AppDataPath 'operator-log.csv'

        # Recreate with headers only
        [PSCustomObject]@{
            Timestamp = $null
            User      = $null
            Action    = $null
            Target    = $null
            Details   = $null
        } | Export-Csv -Path $logFile -NoTypeInformation -Encoding UTF8
        $header = Get-Content -Path $logFile -TotalCount 1
        Set-Content -Path $logFile -Value $header -Encoding UTF8
    }
    catch {
        Write-Verbose "Clear-OperatorLog error: $_"
    }
}

#endregion

#region Alert Thresholds

function Test-TransportAlerts {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$QueueData = @(),

        [Parameter()]
        [hashtable]$DeliveryStats = @{}
    )

    $alerts = [System.Collections.ArrayList]::new()

    try {
        $settings = Get-AppSettings
        $thresholds = $settings.AlertThresholds
        $maxQueueDepth  = if ($thresholds.MaxQueueDepth)   { $thresholds.MaxQueueDepth }   else { 100 }
        $minDeliveryRate = if ($thresholds.MinDeliveryRate) { $thresholds.MinDeliveryRate } else { 95 }

        # Check queue depths
        foreach ($queue in $QueueData) {
            $depth = 0
            if ($queue.PSObject.Properties['MessageCount']) {
                $depth = [int]$queue.MessageCount
            }

            if ($depth -gt $maxQueueDepth) {
                $identity = if ($queue.PSObject.Properties['Identity']) { $queue.Identity } else { 'Unknown' }
                $alerts.Add(@{
                    Level     = 'Warning'
                    Message   = "Queue '$identity' depth ($depth) exceeds threshold ($maxQueueDepth)"
                    Timestamp = Get-Date
                }) | Out-Null
            }

            # Check for Retry delivery type
            $deliveryType = ''
            if ($queue.PSObject.Properties['DeliveryType']) {
                $deliveryType = [string]$queue.DeliveryType
            }
            if ($queue.PSObject.Properties['NextHopDomain']) {
                $nhd = [string]$queue.NextHopDomain
            }
            $status = ''
            if ($queue.PSObject.Properties['Status']) {
                $status = [string]$queue.Status
            }

            if ($deliveryType -match 'Retry' -or $status -match 'Retry') {
                $identity = if ($queue.PSObject.Properties['Identity']) { $queue.Identity } else { 'Unknown' }
                $alerts.Add(@{
                    Level     = 'Warning'
                    Message   = "Retry queue detected: '$identity'"
                    Timestamp = Get-Date
                }) | Out-Null
            }

            # Check for Suspended status
            if ($status -eq 'Suspended') {
                $identity = if ($queue.PSObject.Properties['Identity']) { $queue.Identity } else { 'Unknown' }
                $alerts.Add(@{
                    Level     = 'Info'
                    Message   = "Queue '$identity' is in Suspended status"
                    Timestamp = Get-Date
                }) | Out-Null
            }
        }

        # Check delivery rate
        $delivered = if ($DeliveryStats.ContainsKey('Delivered')) { [int]$DeliveryStats.Delivered } else { 0 }
        $failed    = if ($DeliveryStats.ContainsKey('Failed'))    { [int]$DeliveryStats.Failed }    else { 0 }
        $deferred  = if ($DeliveryStats.ContainsKey('Deferred'))  { [int]$DeliveryStats.Deferred }  else { 0 }
        $total = $delivered + $failed + $deferred

        if ($total -gt 0) {
            $rate = [math]::Round(($delivered / $total) * 100, 2)
            if ($rate -lt $minDeliveryRate) {
                $alerts.Add(@{
                    Level     = 'Critical'
                    Message   = "Delivery rate ($rate%) is below threshold ($minDeliveryRate%)"
                    Timestamp = Get-Date
                }) | Out-Null
            }
        }
    }
    catch {
        Write-Verbose "Test-TransportAlerts error: $_"
    }

    return @($alerts)
}

function Format-AlertText {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$Alerts = @()
    )

    if ($Alerts.Count -eq 0) {
        return 'No active alerts.'
    }

    $lines = [System.Collections.ArrayList]::new()

    foreach ($alert in $Alerts) {
        $indicator = switch ($alert.Level) {
            'Critical' { '[!!!]' }
            'Warning'  { '[!!]'  }
            'Info'     { '[i]'   }
            default    { '[?]'   }
        }

        $ts = if ($alert.Timestamp) {
            ($alert.Timestamp).ToString('HH:mm:ss')
        }
        else {
            (Get-Date).ToString('HH:mm:ss')
        }

        $lines.Add("$indicator $ts - $($alert.Level): $($alert.Message)") | Out-Null
    }

    return ($lines -join "`n")
}

#endregion

#region Statistics

function Get-SenderStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$TrackingData = @()
    )

    try {
        if ($TrackingData.Count -eq 0) { return @() }

        $groups = $TrackingData | Where-Object {
            $_.PSObject.Properties['Sender'] -and $_.Sender
        } | Group-Object -Property Sender

        $results = foreach ($g in $groups) {
            $totalBytes = 0
            $lastSeen = [datetime]::MinValue

            foreach ($entry in $g.Group) {
                if ($entry.PSObject.Properties['TotalBytes']) {
                    $totalBytes += [long]$entry.TotalBytes
                }
                elseif ($entry.PSObject.Properties['MessageSize']) {
                    $totalBytes += [long]$entry.MessageSize
                }

                if ($entry.PSObject.Properties['Timestamp']) {
                    $ts = [datetime]$entry.Timestamp
                    if ($ts -gt $lastSeen) { $lastSeen = $ts }
                }
            }

            [PSCustomObject]@{
                Sender       = $g.Name
                MessageCount = $g.Count
                TotalBytes   = $totalBytes
                LastSeen     = if ($lastSeen -eq [datetime]::MinValue) { $null } else { $lastSeen }
            }
        }

        return @($results | Sort-Object -Property MessageCount -Descending)
    }
    catch {
        Write-Verbose "Get-SenderStatistics error: $_"
        return @()
    }
}

function Get-RecipientStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$TrackingData = @()
    )

    try {
        if ($TrackingData.Count -eq 0) { return @() }

        # Flatten recipients (Recipients field may be an array or semicolon-separated string)
        $expanded = foreach ($entry in $TrackingData) {
            if (-not $entry.PSObject.Properties['Recipients']) { continue }
            $recips = $entry.Recipients
            if ($recips -is [string]) {
                $recips = $recips -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            foreach ($r in $recips) {
                [PSCustomObject]@{
                    Recipient = $r
                    TotalBytes = if ($entry.PSObject.Properties['TotalBytes']) { [long]$entry.TotalBytes }
                                 elseif ($entry.PSObject.Properties['MessageSize']) { [long]$entry.MessageSize }
                                 else { 0 }
                    Timestamp  = if ($entry.PSObject.Properties['Timestamp']) { $entry.Timestamp } else { $null }
                }
            }
        }

        if (-not $expanded) { return @() }

        $groups = @($expanded) | Group-Object -Property Recipient

        $results = foreach ($g in $groups) {
            $totalBytes = ($g.Group | Measure-Object -Property TotalBytes -Sum).Sum
            $lastSeen = ($g.Group | Where-Object { $_.Timestamp } |
                         Sort-Object -Property Timestamp -Descending |
                         Select-Object -First 1).Timestamp

            [PSCustomObject]@{
                Recipient    = $g.Name
                MessageCount = $g.Count
                TotalBytes   = [long]$totalBytes
                LastSeen     = $lastSeen
            }
        }

        return @($results | Sort-Object -Property MessageCount -Descending)
    }
    catch {
        Write-Verbose "Get-RecipientStatistics error: $_"
        return @()
    }
}

function Get-DomainStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$TrackingData = @()
    )

    try {
        if ($TrackingData.Count -eq 0) { return @() }

        $domainData = @{}

        foreach ($entry in $TrackingData) {
            $eventId = ''
            if ($entry.PSObject.Properties['EventId']) {
                $eventId = [string]$entry.EventId
            }

            # Extract sender domain
            if ($entry.PSObject.Properties['Sender'] -and $entry.Sender -match '@(.+)$') {
                $domain = $Matches[1].ToLower()
                if (-not $domainData.ContainsKey($domain)) {
                    $domainData[$domain] = @{ Sent = 0; Received = 0; Failed = 0; Deferred = 0 }
                }

                switch ($eventId) {
                    'SEND'    { $domainData[$domain].Sent++ }
                    'FAIL'    { $domainData[$domain].Failed++ }
                    'DEFER'   { $domainData[$domain].Deferred++ }
                    default   { $domainData[$domain].Sent++ }
                }
            }

            # Extract recipient domains
            if ($entry.PSObject.Properties['Recipients']) {
                $recips = $entry.Recipients
                if ($recips -is [string]) {
                    $recips = $recips -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                }
                foreach ($r in $recips) {
                    if ($r -match '@(.+)$') {
                        $domain = $Matches[1].ToLower()
                        if (-not $domainData.ContainsKey($domain)) {
                            $domainData[$domain] = @{ Sent = 0; Received = 0; Failed = 0; Deferred = 0 }
                        }

                        switch ($eventId) {
                            'DELIVER'  { $domainData[$domain].Received++ }
                            'RECEIVE'  { $domainData[$domain].Received++ }
                            'FAIL'     { $domainData[$domain].Failed++ }
                            'DEFER'    { $domainData[$domain].Deferred++ }
                            default    { $domainData[$domain].Received++ }
                        }
                    }
                }
            }
        }

        $results = foreach ($kvp in $domainData.GetEnumerator()) {
            [PSCustomObject]@{
                Domain        = $kvp.Key
                SentCount     = $kvp.Value.Sent
                ReceivedCount = $kvp.Value.Received
                FailedCount   = $kvp.Value.Failed
                DeferredCount = $kvp.Value.Deferred
            }
        }

        return @($results | Sort-Object -Property { $_.SentCount + $_.ReceivedCount } -Descending)
    }
    catch {
        Write-Verbose "Get-DomainStatistics error: $_"
        return @()
    }
}

function Get-HourlyStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$TrackingData = @()
    )

    try {
        if ($TrackingData.Count -eq 0) { return @() }

        $hourly = @{}

        foreach ($entry in $TrackingData) {
            if (-not $entry.PSObject.Properties['Timestamp']) { continue }

            $hour = ([datetime]$entry.Timestamp).Hour

            if (-not $hourly.ContainsKey($hour)) {
                $hourly[$hour] = @{
                    Received  = 0
                    Delivered = 0
                    Failed    = 0
                    Deferred  = 0
                }
            }

            $eventId = ''
            if ($entry.PSObject.Properties['EventId']) {
                $eventId = [string]$entry.EventId
            }

            switch ($eventId) {
                'RECEIVE'  { $hourly[$hour].Received++ }
                'DELIVER'  { $hourly[$hour].Delivered++ }
                'FAIL'     { $hourly[$hour].Failed++ }
                'DEFER'    { $hourly[$hour].Deferred++ }
                'SEND'     { $hourly[$hour].Delivered++ }
                default    { $hourly[$hour].Received++ }
            }
        }

        $results = foreach ($h in ($hourly.Keys | Sort-Object)) {
            $d = $hourly[$h]
            $total = $d.Delivered + $d.Failed + $d.Deferred
            $rate  = if ($total -gt 0) { [math]::Round(($d.Delivered / $total) * 100, 2) } else { 100.0 }

            [PSCustomObject]@{
                Hour           = $h
                ReceivedCount  = $d.Received
                DeliveredCount = $d.Delivered
                FailedCount    = $d.Failed
                DeferredCount  = $d.Deferred
                DeliveryRate   = $rate
            }
        }

        return @($results)
    }
    catch {
        Write-Verbose "Get-HourlyStatistics error: $_"
        return @()
    }
}

function Get-ConnectorStatistics {
    [CmdletBinding()]
    param(
        [Parameter()]
        [array]$TrackingData = @()
    )

    try {
        if ($TrackingData.Count -eq 0) { return @() }

        $filtered = $TrackingData | Where-Object {
            $_.PSObject.Properties['ConnectorId'] -and $_.ConnectorId
        }

        if (-not $filtered) { return @() }

        $groups = @($filtered) | Group-Object -Property ConnectorId

        $results = foreach ($g in $groups) {
            $uniqueMessages = ($g.Group |
                Where-Object { $_.PSObject.Properties['MessageId'] } |
                Select-Object -ExpandProperty MessageId -Unique).Count

            [PSCustomObject]@{
                Connector      = $g.Name
                EventCount     = $g.Count
                UniqueMessages = $uniqueMessages
            }
        }

        return @($results | Sort-Object -Property EventCount -Descending)
    }
    catch {
        Write-Verbose "Get-ConnectorStatistics error: $_"
        return @()
    }
}

#endregion
