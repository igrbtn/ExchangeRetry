<#
.SYNOPSIS
    Diagnostic functions for Exchange transport: DNS, transport rules,
    certificates, connectivity logs.
.DESCRIPTION
    Dot-source this file to import DNS mail diagnostics (MX/SPF/DKIM/DMARC),
    transport rule inspection, certificate monitoring, and connectivity log parsing.
.NOTES
    Version: 0.1.0
#>

# ═══════════════════════════════════════════════════════════════════════════════
# DNS DIAGNOSTICS
# ═══════════════════════════════════════════════════════════════════════════════

function Get-DnsMxRecord {
    <#
    .SYNOPSIS
        Resolve MX records for a domain.
    .PARAMETER Domain
        The domain to query MX records for.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain
    )

    try {
        $mxRecords = Resolve-DnsName -Name $Domain -Type MX -ErrorAction Stop |
            Where-Object { $_.QueryType -eq 'MX' }

        if (-not $mxRecords) {
            Write-Warning "No MX records found for '$Domain'."
            return @()
        }

        $results = foreach ($mx in $mxRecords | Sort-Object Preference) {
            $ip = $null
            try {
                $aRecord = Resolve-DnsName -Name $mx.NameExchange -Type A -ErrorAction Stop |
                    Where-Object { $_.QueryType -eq 'A' } | Select-Object -First 1
                $ip = $aRecord.IPAddress
            }
            catch {
                Write-Verbose "Could not resolve A record for $($mx.NameExchange): $_"
            }

            [PSCustomObject]@{
                Domain     = $Domain
                Priority   = $mx.Preference
                MailServer = $mx.NameExchange
                IP         = $ip
            }
        }
        return $results
    }
    catch {
        Write-Warning "Failed to resolve MX for '$Domain': $_"
        return @()
    }
}

function Get-DnsSpfRecord {
    <#
    .SYNOPSIS
        Get and parse SPF TXT record for a domain.
    .PARAMETER Domain
        The domain to query SPF for.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain
    )

    try {
        $txtRecords = Resolve-DnsName -Name $Domain -Type TXT -ErrorAction Stop |
            Where-Object { $_.QueryType -eq 'TXT' }

        $spfRecord = $txtRecords | Where-Object {
            $_.Strings -join '' -match '^v=spf1'
        } | Select-Object -First 1

        if (-not $spfRecord) {
            Write-Warning "No SPF record found for '$Domain'."
            return [PSCustomObject]@{
                Domain     = $Domain
                Record     = $null
                Mechanisms = @()
                Qualifier  = $null
            }
        }

        $raw = ($spfRecord.Strings -join '')
        $tokens = $raw -split '\s+' | Where-Object { $_ -and $_ -ne 'v=spf1' }

        $mechanisms = @()
        $allQualifier = $null

        foreach ($token in $tokens) {
            $qualifier = '+'
            $mechanism = $token

            if ($token -match '^([+\-~?])(.+)$') {
                $qualifier = $Matches[1]
                $mechanism = $Matches[2]
            }

            if ($mechanism -match '^all$') {
                $allQualifier = $qualifier
            }

            $mechanisms += [PSCustomObject]@{
                Qualifier = $qualifier
                Mechanism = $mechanism
                Raw       = $token
            }
        }

        return [PSCustomObject]@{
            Domain     = $Domain
            Record     = $raw
            Mechanisms = $mechanisms
            Qualifier  = $allQualifier
        }
    }
    catch {
        Write-Warning "Failed to query SPF for '$Domain': $_"
        return [PSCustomObject]@{
            Domain     = $Domain
            Record     = $null
            Mechanisms = @()
            Qualifier  = $null
        }
    }
}

function Get-DnsDkimRecord {
    <#
    .SYNOPSIS
        Get DKIM TXT record for a domain.
    .PARAMETER Domain
        The domain to query DKIM for.
    .PARAMETER Selector
        DKIM selector (default: tries common selectors).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain,

        [string]$Selector = 'default'
    )

    $selectors = if ($Selector -ne 'default') {
        @($Selector)
    }
    else {
        @('default', 'selector1', 'selector2', 'google', 'k1')
    }

    foreach ($sel in $selectors) {
        $dkimName = "$sel._domainkey.$Domain"
        try {
            $txtRecords = Resolve-DnsName -Name $dkimName -Type TXT -ErrorAction Stop |
                Where-Object { $_.QueryType -eq 'TXT' }

            if (-not $txtRecords) { continue }

            $raw = ($txtRecords[0].Strings -join '')

            $keyType = $null
            if ($raw -match 'k=([^;]+)') { $keyType = $Matches[1].Trim() }

            $keyData = $null
            if ($raw -match 'p=([^;]+)') { $keyData = $Matches[1].Trim() }

            return [PSCustomObject]@{
                Domain   = $Domain
                Selector = $sel
                Record   = $raw
                KeyType  = $keyType
                KeyData  = $keyData
            }
        }
        catch {
            Write-Verbose "No DKIM record at '$dkimName': $_"
        }
    }

    Write-Warning "No DKIM record found for '$Domain' (tried selectors: $($selectors -join ', '))."
    return [PSCustomObject]@{
        Domain   = $Domain
        Selector = $null
        Record   = $null
        KeyType  = $null
        KeyData  = $null
    }
}

function Get-DnsDmarcRecord {
    <#
    .SYNOPSIS
        Get and parse _dmarc TXT record for a domain.
    .PARAMETER Domain
        The domain to query DMARC for.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain
    )

    try {
        $dmarcName = "_dmarc.$Domain"
        $txtRecords = Resolve-DnsName -Name $dmarcName -Type TXT -ErrorAction Stop |
            Where-Object { $_.QueryType -eq 'TXT' }

        $dmarcRecord = $txtRecords | Where-Object {
            $_.Strings -join '' -match '^v=DMARC1'
        } | Select-Object -First 1

        if (-not $dmarcRecord) {
            Write-Warning "No DMARC record found for '$Domain'."
            return [PSCustomObject]@{
                Domain          = $Domain
                Record          = $null
                Policy          = $null
                SubdomainPolicy = $null
                Percentage      = $null
                ReportUri       = $null
                ForensicUri     = $null
                DkimAlignment   = $null
                SpfAlignment    = $null
            }
        }

        $raw = ($dmarcRecord.Strings -join '')

        $policy          = if ($raw -match '\bp=([^;]+)')    { $Matches[1].Trim() } else { $null }
        $subPolicy       = if ($raw -match '\bsp=([^;]+)')   { $Matches[1].Trim() } else { $null }
        $pct             = if ($raw -match '\bpct=([^;]+)')  { [int]$Matches[1].Trim() } else { $null }
        $rua             = if ($raw -match '\brua=([^;]+)')  { $Matches[1].Trim() } else { $null }
        $ruf             = if ($raw -match '\bruf=([^;]+)')  { $Matches[1].Trim() } else { $null }
        $adkim           = if ($raw -match '\badkim=([^;]+)') { $Matches[1].Trim() } else { $null }
        $aspf            = if ($raw -match '\baspf=([^;]+)')  { $Matches[1].Trim() } else { $null }

        return [PSCustomObject]@{
            Domain          = $Domain
            Record          = $raw
            Policy          = $policy
            SubdomainPolicy = $subPolicy
            Percentage      = $pct
            ReportUri       = $rua
            ForensicUri     = $ruf
            DkimAlignment   = $adkim
            SpfAlignment    = $aspf
        }
    }
    catch {
        Write-Warning "Failed to query DMARC for '$Domain': $_"
        return [PSCustomObject]@{
            Domain          = $Domain
            Record          = $null
            Policy          = $null
            SubdomainPolicy = $null
            Percentage      = $null
            ReportUri       = $null
            ForensicUri     = $null
            DkimAlignment   = $null
            SpfAlignment    = $null
        }
    }
}

function Test-DomainMailHealth {
    <#
    .SYNOPSIS
        Run all DNS mail checks (MX, SPF, DKIM, DMARC) and return a health assessment.
    .PARAMETER Domain
        The domain to check.
    .PARAMETER DkimSelector
        DKIM selector to use (default: tries common selectors).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Domain,

        [string]$DkimSelector = 'default'
    )

    $mx    = Get-DnsMxRecord    -Domain $Domain
    $spf   = Get-DnsSpfRecord   -Domain $Domain
    $dkim  = Get-DnsDkimRecord  -Domain $Domain -Selector $DkimSelector
    $dmarc = Get-DnsDmarcRecord -Domain $Domain

    # --- Health assessment ---
    $issues = @()

    $hasMx = ($mx | Measure-Object).Count -gt 0
    if (-not $hasMx) { $issues += 'No MX records' }

    $hasSpf = [bool]$spf.Record
    if (-not $hasSpf) {
        $issues += 'No SPF record'
    }
    elseif ($spf.Qualifier -eq '+') {
        $issues += 'SPF too permissive (+all)'
    }

    $hasDkim = [bool]$dkim.Record
    if (-not $hasDkim) { $issues += 'No DKIM record found' }

    $hasDmarc = [bool]$dmarc.Record
    if (-not $hasDmarc) {
        $issues += 'No DMARC record'
    }
    elseif ($dmarc.Policy -eq 'none') {
        $issues += 'DMARC policy is none (monitor only)'
    }

    $health = if ($issues.Count -eq 0) {
        'Good'
    }
    elseif (-not $hasMx -or ($hasSpf -and $spf.Qualifier -eq '+') -or (-not $hasDmarc)) {
        'Bad'
    }
    else {
        'Warning'
    }

    return [PSCustomObject]@{
        Domain     = $Domain
        MX         = $mx
        SPF        = $spf
        DKIM       = $dkim
        DMARC      = $dmarc
        Health     = $health
        Issues     = $issues
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# TRANSPORT RULES
# ═══════════════════════════════════════════════════════════════════════════════

function Get-ExchangeTransportRules {
    <#
    .SYNOPSIS
        Retrieve Exchange transport rules with formatted details.
    #>
    [CmdletBinding()]
    param()

    try {
        $rules = Get-TransportRule -ErrorAction Stop

        if (-not $rules) {
            Write-Warning "No transport rules found."
            return @()
        }

        $results = foreach ($rule in $rules) {
            # Format conditions
            $conditions = @()
            if ($rule.SentTo)                { $conditions += "SentTo: $($rule.SentTo -join ', ')" }
            if ($rule.SentToScope)           { $conditions += "SentToScope: $($rule.SentToScope)" }
            if ($rule.From)                  { $conditions += "From: $($rule.From -join ', ')" }
            if ($rule.FromScope)             { $conditions += "FromScope: $($rule.FromScope)" }
            if ($rule.SubjectContainsWords)  { $conditions += "SubjectContains: $($rule.SubjectContainsWords -join ', ')" }
            if ($rule.FromMemberOf)          { $conditions += "FromMemberOf: $($rule.FromMemberOf -join ', ')" }
            if ($rule.SentToMemberOf)        { $conditions += "SentToMemberOf: $($rule.SentToMemberOf -join ', ')" }
            if ($rule.HeaderContainsMessageHeader) {
                $conditions += "HeaderContains: $($rule.HeaderContainsMessageHeader)=$($rule.HeaderContainsWords -join ', ')"
            }

            # Format actions
            $actions = @()
            if ($rule.RejectMessageReasonText)   { $actions += "Reject: $($rule.RejectMessageReasonText)" }
            if ($rule.AddToRecipients)           { $actions += "AddRecipient: $($rule.AddToRecipients -join ', ')" }
            if ($rule.CopyTo)                    { $actions += "CopyTo: $($rule.CopyTo -join ', ')" }
            if ($rule.BlindCopyTo)               { $actions += "BccTo: $($rule.BlindCopyTo -join ', ')" }
            if ($rule.RedirectMessageTo)         { $actions += "Redirect: $($rule.RedirectMessageTo -join ', ')" }
            if ($rule.PrependSubject)             { $actions += "PrependSubject: $($rule.PrependSubject)" }
            if ($rule.SetHeaderName)             { $actions += "SetHeader: $($rule.SetHeaderName)=$($rule.SetHeaderValue)" }
            if ($rule.ApplyClassification)       { $actions += "Classification: $($rule.ApplyClassification)" }
            if ($rule.SetSCL -ne $null)          { $actions += "SetSCL: $($rule.SetSCL)" }
            if ($rule.DeleteMessage)             { $actions += "DeleteMessage" }

            # Format exceptions
            $exceptions = @()
            if ($rule.ExceptIfSentTo)           { $exceptions += "ExceptSentTo: $($rule.ExceptIfSentTo -join ', ')" }
            if ($rule.ExceptIfFrom)             { $exceptions += "ExceptFrom: $($rule.ExceptIfFrom -join ', ')" }
            if ($rule.ExceptIfSubjectContainsWords) {
                $exceptions += "ExceptSubject: $($rule.ExceptIfSubjectContainsWords -join ', ')"
            }

            [PSCustomObject]@{
                Name                = $rule.Name
                State               = $rule.State
                Priority            = $rule.Priority
                Mode                = $rule.Mode
                SentTo              = $rule.SentTo
                SentToScope         = $rule.SentToScope
                From                = $rule.From
                FromScope           = $rule.FromScope
                SubjectContainsWords = $rule.SubjectContainsWords
                Actions             = ($actions -join '; ')
                Conditions          = ($conditions -join '; ')
                Exceptions          = ($exceptions -join '; ')
            }
        }
        return $results
    }
    catch {
        Write-Warning "Failed to get transport rules: $_"
        return @()
    }
}

function Get-TransportRuleReport {
    <#
    .SYNOPSIS
        Generate a text report of Exchange transport rules.
    #>
    [CmdletBinding()]
    param()

    $rules = Get-ExchangeTransportRules
    $sb = [System.Text.StringBuilder]::new()

    [void]$sb.AppendLine("=" * 70)
    [void]$sb.AppendLine("TRANSPORT RULES REPORT")
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("=" * 70)

    if (-not $rules -or ($rules | Measure-Object).Count -eq 0) {
        [void]$sb.AppendLine("`nNo transport rules found.")
        return $sb.ToString()
    }

    $total    = ($rules | Measure-Object).Count
    $enabled  = ($rules | Where-Object { $_.State -eq 'Enabled' } | Measure-Object).Count
    $disabled = $total - $enabled

    [void]$sb.AppendLine("`nSummary:")
    [void]$sb.AppendLine("  Total rules : $total")
    [void]$sb.AppendLine("  Enabled     : $enabled")
    [void]$sb.AppendLine("  Disabled    : $disabled")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("-" * 70)

    foreach ($rule in $rules) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("Rule: $($rule.Name)")
        [void]$sb.AppendLine("  State      : $($rule.State)")
        [void]$sb.AppendLine("  Priority   : $($rule.Priority)")
        [void]$sb.AppendLine("  Mode       : $($rule.Mode)")

        if ($rule.Conditions) {
            [void]$sb.AppendLine("  Conditions : $($rule.Conditions)")
        }
        if ($rule.Actions) {
            [void]$sb.AppendLine("  Actions    : $($rule.Actions)")
        }
        if ($rule.Exceptions) {
            [void]$sb.AppendLine("  Exceptions : $($rule.Exceptions)")
        }
        [void]$sb.AppendLine("  " + ("-" * 66))
    }

    return $sb.ToString()
}

# ═══════════════════════════════════════════════════════════════════════════════
# CERTIFICATE MONITORING
# ═══════════════════════════════════════════════════════════════════════════════

function Get-ExchangeCertificates {
    <#
    .SYNOPSIS
        Get Exchange certificates with expiry status.
    .PARAMETER ExpiryWarningDays
        Number of days before expiry to flag as 'Expiring' (default: 30).
    #>
    [CmdletBinding()]
    param(
        [int]$ExpiryWarningDays = 30
    )

    try {
        $certs = Get-ExchangeCertificate -ErrorAction Stop

        if (-not $certs) {
            Write-Warning "No Exchange certificates found."
            return @()
        }

        $now = Get-Date
        $results = foreach ($cert in $certs) {
            $daysUntil = ($cert.NotAfter - $now).Days
            $status = if ($daysUntil -lt 0) {
                'Expired'
            }
            elseif ($daysUntil -le $ExpiryWarningDays) {
                'Expiring'
            }
            else {
                'Valid'
            }

            [PSCustomObject]@{
                Thumbprint     = $cert.Thumbprint
                Subject        = $cert.Subject
                Issuer         = $cert.Issuer
                NotBefore      = $cert.NotBefore
                NotAfter       = $cert.NotAfter
                DaysUntilExpiry = $daysUntil
                Status         = $status
                Services       = $cert.Services
                IsSelfSigned   = $cert.IsSelfSigned
            }
        }
        return $results
    }
    catch {
        Write-Warning "Failed to get Exchange certificates: $_"
        return @()
    }
}

function Get-ConnectorCertificateBindings {
    <#
    .SYNOPSIS
        Show certificate bindings for Send and Receive connectors.
    #>
    [CmdletBinding()]
    param()

    try {
        $certs = Get-ExchangeCertificate -ErrorAction Stop
        $certLookup = @{}
        foreach ($cert in $certs) {
            $certLookup[$cert.Thumbprint] = $cert
        }
    }
    catch {
        Write-Warning "Failed to get certificates: $_"
        $certLookup = @{}
    }

    $bindings = @()

    # Send connectors
    try {
        $sendConnectors = Get-SendConnector -ErrorAction Stop
        foreach ($conn in $sendConnectors) {
            $tlsCertName = $conn.TlsCertificateName
            $matchedCert = $null

            if ($tlsCertName) {
                # TlsCertificateName format: <I>issuer<S>subject
                $issuer  = $null
                $subject = $null
                if ($tlsCertName -match '<I>(.+)<S>(.+)') {
                    $issuer  = $Matches[1]
                    $subject = $Matches[2]
                }

                foreach ($cert in $certs) {
                    if ($cert.Issuer -eq $issuer -and $cert.Subject -eq $subject) {
                        $matchedCert = $cert
                        break
                    }
                }
            }

            $bindings += [PSCustomObject]@{
                ConnectorType    = 'Send'
                ConnectorName    = $conn.Name
                TlsCertName      = $tlsCertName
                CertThumbprint   = if ($matchedCert) { $matchedCert.Thumbprint } else { $null }
                CertSubject      = if ($matchedCert) { $matchedCert.Subject }    else { $null }
                CertExpiry       = if ($matchedCert) { $matchedCert.NotAfter }   else { $null }
                CertStatus       = if ($matchedCert) {
                    $days = ($matchedCert.NotAfter - (Get-Date)).Days
                    if ($days -lt 0) { 'Expired' } elseif ($days -le 30) { 'Expiring' } else { 'Valid' }
                } else { 'Unknown' }
            }
        }
    }
    catch {
        Write-Warning "Failed to get send connectors: $_"
    }

    # Receive connectors
    try {
        $recvConnectors = Get-ReceiveConnector -ErrorAction Stop
        foreach ($conn in $recvConnectors) {
            $tlsCertName = $conn.TlsCertificateName
            $matchedCert = $null

            if ($tlsCertName) {
                $issuer  = $null
                $subject = $null
                if ($tlsCertName -match '<I>(.+)<S>(.+)') {
                    $issuer  = $Matches[1]
                    $subject = $Matches[2]
                }

                foreach ($cert in $certs) {
                    if ($cert.Issuer -eq $issuer -and $cert.Subject -eq $subject) {
                        $matchedCert = $cert
                        break
                    }
                }
            }

            $bindings += [PSCustomObject]@{
                ConnectorType    = 'Receive'
                ConnectorName    = $conn.Identity
                TlsCertName      = $tlsCertName
                CertThumbprint   = if ($matchedCert) { $matchedCert.Thumbprint } else { $null }
                CertSubject      = if ($matchedCert) { $matchedCert.Subject }    else { $null }
                CertExpiry       = if ($matchedCert) { $matchedCert.NotAfter }   else { $null }
                CertStatus       = if ($matchedCert) {
                    $days = ($matchedCert.NotAfter - (Get-Date)).Days
                    if ($days -lt 0) { 'Expired' } elseif ($days -le 30) { 'Expiring' } else { 'Valid' }
                } else { 'Unknown' }
            }
        }
    }
    catch {
        Write-Warning "Failed to get receive connectors: $_"
    }

    return $bindings
}

function Get-CertificateReport {
    <#
    .SYNOPSIS
        Generate a text report of Exchange certificates and connector bindings.
    #>
    [CmdletBinding()]
    param(
        [int]$ExpiryWarningDays = 30
    )

    $certs    = Get-ExchangeCertificates -ExpiryWarningDays $ExpiryWarningDays
    $bindings = Get-ConnectorCertificateBindings

    $sb = [System.Text.StringBuilder]::new()

    [void]$sb.AppendLine("=" * 70)
    [void]$sb.AppendLine("EXCHANGE CERTIFICATE REPORT")
    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("=" * 70)

    # --- Certificates ---
    [void]$sb.AppendLine("`nCERTIFICATES")
    [void]$sb.AppendLine("-" * 70)

    if (-not $certs -or ($certs | Measure-Object).Count -eq 0) {
        [void]$sb.AppendLine("  No certificates found.")
    }
    else {
        $expired  = ($certs | Where-Object { $_.Status -eq 'Expired' }  | Measure-Object).Count
        $expiring = ($certs | Where-Object { $_.Status -eq 'Expiring' } | Measure-Object).Count
        $valid    = ($certs | Where-Object { $_.Status -eq 'Valid' }    | Measure-Object).Count

        [void]$sb.AppendLine("  Total: $(($certs | Measure-Object).Count)  |  Valid: $valid  |  Expiring: $expiring  |  Expired: $expired")
        [void]$sb.AppendLine("")

        foreach ($cert in $certs) {
            $flag = switch ($cert.Status) {
                'Expired'  { '[!!!]' }
                'Expiring' { '[!]  ' }
                default    { '     ' }
            }
            [void]$sb.AppendLine("  $flag $($cert.Subject)")
            [void]$sb.AppendLine("        Thumbprint : $($cert.Thumbprint)")
            [void]$sb.AppendLine("        Issuer     : $($cert.Issuer)")
            [void]$sb.AppendLine("        Valid      : $($cert.NotBefore.ToString('yyyy-MM-dd')) - $($cert.NotAfter.ToString('yyyy-MM-dd'))  ($($cert.DaysUntilExpiry) days left)")
            [void]$sb.AppendLine("        Services   : $($cert.Services)")
            [void]$sb.AppendLine("        SelfSigned : $($cert.IsSelfSigned)")
            [void]$sb.AppendLine("")
        }
    }

    # --- Connector bindings ---
    [void]$sb.AppendLine("CONNECTOR CERTIFICATE BINDINGS")
    [void]$sb.AppendLine("-" * 70)

    if (-not $bindings -or ($bindings | Measure-Object).Count -eq 0) {
        [void]$sb.AppendLine("  No connector bindings found.")
    }
    else {
        foreach ($b in $bindings) {
            $flag = switch ($b.CertStatus) {
                'Expired'  { '[!!!]' }
                'Expiring' { '[!]  ' }
                'Unknown'  { '[?]  ' }
                default    { '     ' }
            }
            [void]$sb.AppendLine("  $flag [$($b.ConnectorType)] $($b.ConnectorName)")
            if ($b.CertThumbprint) {
                [void]$sb.AppendLine("        Certificate : $($b.CertSubject) (expires $($b.CertExpiry.ToString('yyyy-MM-dd')))")
            }
            elseif ($b.TlsCertName) {
                [void]$sb.AppendLine("        TlsCertName : $($b.TlsCertName)  (certificate NOT matched)")
            }
            else {
                [void]$sb.AppendLine("        Certificate : (none configured)")
            }
        }
    }

    # --- Warnings ---
    $warnings = @()
    if ($certs) {
        $certs | Where-Object { $_.Status -eq 'Expired' }  | ForEach-Object { $warnings += "EXPIRED: $($_.Subject) expired on $($_.NotAfter.ToString('yyyy-MM-dd'))" }
        $certs | Where-Object { $_.Status -eq 'Expiring' } | ForEach-Object { $warnings += "EXPIRING: $($_.Subject) expires in $($_.DaysUntilExpiry) days" }
    }
    if ($bindings) {
        $bindings | Where-Object { $_.CertStatus -eq 'Expired' }  | ForEach-Object { $warnings += "CONNECTOR '$($_.ConnectorName)' uses an EXPIRED certificate" }
        $bindings | Where-Object { $_.CertStatus -eq 'Expiring' } | ForEach-Object { $warnings += "CONNECTOR '$($_.ConnectorName)' uses an EXPIRING certificate" }
        $bindings | Where-Object { $_.CertStatus -eq 'Unknown' -and $_.TlsCertName } | ForEach-Object {
            $warnings += "CONNECTOR '$($_.ConnectorName)' references a certificate that could not be matched"
        }
    }

    if ($warnings.Count -gt 0) {
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("WARNINGS")
        [void]$sb.AppendLine("-" * 70)
        foreach ($w in $warnings) {
            [void]$sb.AppendLine("  * $w")
        }
    }

    return $sb.ToString()
}

# ═══════════════════════════════════════════════════════════════════════════════
# CONNECTIVITY LOG PARSING
# ═══════════════════════════════════════════════════════════════════════════════

function Parse-ConnectivityLog {
    <#
    .SYNOPSIS
        Parse Exchange connectivity log files (CSV with #Fields: header).
    .PARAMETER LogPath
        Path to log file or directory containing log files.
    .PARAMETER Filter
        Optional filter string to match against log entries.
    .PARAMETER MaxFiles
        Maximum number of log files to parse (default: 50).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LogPath,

        [string]$Filter,

        [int]$MaxFiles = 50
    )

    $results = @()

    try {
        # Determine if path is a file or directory
        if (Test-Path -Path $LogPath -PathType Container) {
            $logFiles = Get-ChildItem -Path $LogPath -Filter '*.log' -File |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $MaxFiles
        }
        elseif (Test-Path -Path $LogPath -PathType Leaf) {
            $logFiles = @(Get-Item -Path $LogPath)
        }
        else {
            Write-Warning "Path not found: '$LogPath'"
            return @()
        }

        if (-not $logFiles -or $logFiles.Count -eq 0) {
            Write-Warning "No log files found at '$LogPath'."
            return @()
        }

        foreach ($file in $logFiles) {
            $headers = $null
            $lineNum = 0

            try {
                foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {
                    $lineNum++

                    # Skip comment lines, but capture the #Fields: header
                    if ($line.StartsWith('#')) {
                        if ($line -match '^#Fields:\s*(.+)$') {
                            $headers = $Matches[1] -split ',' | ForEach-Object { $_.Trim() }
                        }
                        continue
                    }

                    # Skip empty lines
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }

                    # Apply filter if specified
                    if ($Filter -and $line -notmatch [regex]::Escape($Filter)) { continue }

                    if (-not $headers) {
                        # Fallback: use default connectivity log columns
                        $headers = @('date-time', 'session', 'source', 'Destination', 'direction', 'description')
                    }

                    $fields = $line -split ','
                    $entry = [ordered]@{
                        SourceFile = $file.Name
                    }

                    for ($i = 0; $i -lt $headers.Count; $i++) {
                        $value = if ($i -lt $fields.Count) { $fields[$i].Trim() } else { $null }
                        $entry[$headers[$i]] = $value
                    }

                    $results += [PSCustomObject]$entry
                }
            }
            catch {
                Write-Warning "Error reading '$($file.FullName)' at line $lineNum`: $_"
            }
        }

        return $results
    }
    catch {
        Write-Warning "Failed to parse connectivity logs: $_"
        return @()
    }
}
