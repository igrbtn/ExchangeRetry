# ExchangeRetry

Full-featured PowerShell GUI + CLI for monitoring and managing Microsoft Exchange transport pipeline. All operations run asynchronously — the GUI never freezes.

## Features

| Tab | Description |
|-----|-------------|
| **Dashboard** | Transport health at a glance: queue status, delivery rates (1h/24h), connector health, alerts |
| **Queues** | Queue management with retry/suspend/remove, filtering, auto-refresh, and **recent errors panel** (FAIL/DEFER/DSN) |
| **Message Tracking** | Search `Get-MessageTrackingLog` with filters: EventId, Sender, Recipient, Subject, date range. **Show Message Path** and **Cross-Server Trace** |
| **Protocol Logs** | Parse SMTP Send/Receive protocol log files (Exchange CSV format) with text filtering |
| **Log Search** | Full-text search across any transport log files with context display |
| **Header Analyzer** | Parse email headers: Received hops with delays, SPF/DKIM/DMARC, TLS detection, X-MS-Exchange-* headers |
| **Diagnostics** | DNS mail health (MX/SPF/DKIM/DMARC), Transport Rules viewer, Certificate manager, Connectivity Logs |
| **Statistics** | Mail flow analytics: by Sender, Recipient, Domain, Hour, Connector |
| **Reports** | 9 report types: Full, Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Pipeline, BackPressure |

**Server Scope** — auto-discovers all Exchange transport servers on connect. Query a specific server or all at once.

**Async Architecture** — all Exchange operations run in PowerShell runspaces with progress bars and a collapsible job console at the bottom.

**Export** — all tabs support export to CSV/JSON.

## Transport Pipeline Coverage

```
Internet/Internal
      |
      v
+-----------------+     +-----------------+
| Receive         |---->| Protocol Logs   |  Tab: Protocol Logs
| Connector       |     | (SMTP Receive)  |
+-----------------+     +-----------------+
      |
      v
+-----------------+     +-----------------+
| Transport       |---->| Agent Log       |  Tab: Reports > AgentLog
| Agents          |     +-----------------+
+-----------------+
      |
      v
+-----------------+     +-----------------+
| Categorizer     |---->| Message Tracking|  Tab: Message Tracking
| (routing)       |     | RECEIVE/SUBMIT  |
+-----------------+     +-----------------+
      |
      v
+-----------------+     +-----------------+
| Queue           |---->| Queue Viewer    |  Tab: Queues + Errors
| (delivery)      |     | + FAIL/DEFER    |
+-----------------+     +-----------------+
      |
      v
+-----------------+     +-----------------+
| Send            |---->| Protocol Logs   |  Tab: Protocol Logs
| Connector       |     | (SMTP Send)     |
+-----------------+     +-----------------+
      |
      v
+-----------------+     +-----------------+
| Delivery/Relay  |---->| Message Tracking|  Tab: Message Tracking
|                 |     | DELIVER/SEND    |  "Show Message Path"
+-----------------+     +-----------------+

+-----------------+
| Email Headers   |     Tab: Header Analyzer
| (hops, auth,    |     Parse & visualize
|  X-Headers)     |
+-----------------+

+-----------------+
| DNS Records     |     Tab: Diagnostics > DNS
| MX/SPF/DKIM/   |     Domain mail health check
| DMARC           |
+-----------------+

+-----------------+
| Transport Logs  |     Tab: Log Search
| (connectivity,  |     Text search + context
|  routing, etc)  |
+-----------------+
```

## Quick Start

### Requirements
- Windows PowerShell 5.1+ or PowerShell 7+
- Exchange Management Shell (EMS) or remote access to Exchange server
- Kerberos authentication to Exchange

### GUI
```powershell
.\ExchangeRetry.ps1
```
1. Enter Exchange server FQDN and click **Connect**
2. All transport servers are auto-discovered in the **Scope** dropdown
3. Navigate tabs to monitor and manage transport
4. Watch the **Job Console** at the bottom for operation progress

### CLI
```powershell
# Parse email headers from file
.\ExchangeTrace.ps1 -HeaderFile .\header.txt

# Track message by Message-ID
.\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01

# Search by sender (last 7 days)
.\ExchangeTrace.ps1 -Sender user@domain.com -StartDate (Get-Date).AddDays(-7) -Server exchange01

# Search transport log files
.\ExchangeTrace.ps1 -TransportLogPath "\\exchange01\TransportLogs" -SearchPattern "user@domain.com"

# Full transport report
.\ExchangeTrace.ps1 -Report Full -Server exchange01

# Delivery summary
.\ExchangeTrace.ps1 -Report Summary -Server exchange01

# Export to CSV
.\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01 -OutputFile results.csv -OutputFormat CSV
```

### Environment Variable
```powershell
$env:EXCHANGE_SERVER = "exchange01.domain.local"
.\ExchangeRetry.ps1  # auto-fills server field
```

## Async Architecture

```
┌──────────────────────────────────────────────────┐
│ WinForms UI Thread                                │
│  ┌─────────┐  ┌─────────┐  ┌─────────┐          │
│  │ Button   │  │ Timer   │  │ Poller  │ (200ms)  │
│  │ Click    │  │ Auto-   │  │ checks  │          │
│  │          │  │ refresh │  │ jobs    │          │
│  └────┬─────┘  └────┬────┘  └────┬────┘          │
│       │              │            │               │
│       v              v            v               │
│  Start-AsyncJob    Start-     Update-AsyncJobs    │
│  (creates          AsyncJob   (EndInvoke,         │
│   runspace)                    callbacks)          │
└───────┬──────────────┬───────────┬────────────────┘
        │              │           │
        v              v           v
┌───────────┐  ┌───────────┐  ┌─────────────────┐
│ Runspace  │  │ Runspace  │  │ Job Console     │
│ (Exchange │  │ (Exchange │  │ [12:30] START #1│
│  cmdlets) │  │  cmdlets) │  │ [12:31] DONE #1│
└───────────┘  └───────────┘  └─────────────────┘
```

- Every Exchange operation (Connect, Queues, Tracking, Reports, DNS, Certs, Stats) runs in its own runspace
- UI thread stays responsive — never blocks on network calls
- Progress bar shows when jobs are running
- Auto-refresh skips if jobs are already running (prevents pileup)
- All callbacks wrapped in try/catch — core never crashes from module errors

## Project Structure

```
ExchangeRetry/
├── ExchangeRetry.ps1           # GUI (WinForms) — 9 tabs, async, job console
├── ExchangeTrace.ps1           # CLI — same functions, console output
├── lib/
│   ├── Core.ps1                # Exchange functions (connection, queues, tracking, reports)
│   ├── Diagnostics.ps1         # DNS, transport rules, certificates, connectivity logs
│   ├── Monitoring.ps1          # Settings, cache, operator log, alerts, statistics
│   └── AsyncRunner.ps1         # Async framework (runspaces, progress, job console)
├── docs/
│   ├── architecture.md         # GUI layout diagrams, tab structure
│   └── data-flow.md            # Data flow diagrams for all features
├── tests/
│   ├── ExchangeRetry.Tests.ps1 # Pester tests for GUI functions
│   └── ExchangeTrace.Tests.ps1 # Pester tests for CLI functions
├── CLAUDE.md                   # Dev instructions
├── .env.example                # Environment variables template
└── .gitignore
```

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **F5** | Refresh current tab |
| **Ctrl+E** | Export current tab data |
| **Ctrl+F** | Focus search/filter field |
| **Escape** | Clear search/filter |

## Testing

```powershell
Invoke-Pester -Path ./tests/
```

## Version

**0.5.0** — Full async architecture with runspaces, job console, progress bars. 9 tabs: Dashboard, Queues+Errors, Message Tracking with Path/Cross-Server, Protocol Logs, Log Search, Header Analyzer, Diagnostics (DNS/Rules/Certs/ConnLogs), Statistics, Reports. Comprehensive error handling — core never crashes.
