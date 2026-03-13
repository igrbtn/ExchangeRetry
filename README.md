# ExchangeRetry

Full-featured PowerShell GUI + CLI for monitoring and managing Microsoft Exchange transport pipeline.

## Features

| Tab | Description |
|-----|-------------|
| **Dashboard** | Transport health at a glance: queue status, delivery rates (1h/24h), connector health, recent failures |
| **Queues** | Queue management with retry/suspend/remove, filtering, auto-refresh, and **recent errors panel** (FAIL/DEFER/DSN) |
| **Message Tracking** | Search `Get-MessageTrackingLog` with filters: EventId, Sender, Recipient, Subject, date range. **Show Message Path** — step-by-step route visualization |
| **Protocol Logs** | Parse SMTP Send/Receive protocol log files (Exchange CSV format) with text filtering |
| **Log Search** | Full-text search across any transport log files with context display |
| **Header Analyzer** | Parse email headers: Received hops with delays, SPF/DKIM/DMARC, TLS detection, X-MS-Exchange-* headers |
| **Reports** | 9 report types: Full, Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Pipeline, BackPressure |

**Server Scope** — auto-discovers all Exchange transport servers on connect. Query a specific server or all at once.

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

## GUI Screenshots Layout

### Queues Tab (3-panel)
```
+------------------------------------------+
| Queues grid (with filters, auto-refresh) |
+------------------------------------------+
| Messages in selected queue               |
| [Retry] [Suspend] [Remove] [NDR]        |
+------------------------------------------+
| Recent Errors (FAIL/DEFER/DSN - 24h)     |
| FAIL: 3 | DEFER: 7 | DSN: 1 | Total: 11|
+------------------------------------------+
```

### Message Tracking
```
+------------------------------------------+
| Filters: MsgID, Sender, Recipient,       |
|          Subject, EventId, Date range    |
+------------------------------------------+
| Tracking results grid                    |
| [Export] [Show Message Path]  N event(s) |
+------------------------------------------+
```

## Reports

| Report | Exchange Cmdlets Used |
|--------|----------------------|
| Queues | `Get-Queue` |
| Connectors | `Get-SendConnector`, `Get-ReceiveConnector` |
| AgentLog | `Get-TransportAgent`, `Get-AgentLog` |
| RoutingTable | `Get-TransportConfig`, `Get-TransportService` |
| DSN | `Get-TransportConfig`, `Get-MessageTrackingLog -EventId DSN` |
| Summary | `Get-MessageTrackingLog` (full 24h analysis) |
| Pipeline | `Get-TransportAgent` (sorted by priority) |
| BackPressure | `Get-ExchangeDiagnosticInfo -Component ResourceThrottling` |
| Full | All of the above |

## Project Structure

```
ExchangeRetry/
├── ExchangeRetry.ps1           # GUI (WinForms) — 7 tabs, full transport monitoring
├── ExchangeTrace.ps1           # CLI — same functions, console output
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

## Testing

```powershell
Invoke-Pester -Path ./tests/
```

## Version

**0.3.0** — Full transport monitoring with Dashboard, Queues+Errors, Message Tracking with Path visualization, Protocol Log parser, Log Search, Header Analyzer, 9 report types, server scope selector.
