# ExchangeRetry

Full-featured PowerShell GUI + CLI for monitoring and managing Microsoft Exchange transport pipeline. Auto-detects Exchange Management Shell, auto-discovers servers, and runs all operations asynchronously - the GUI never freezes.

## Features

| Tab | Description |
|-----|-------------|
| **Dashboard** | Transport health at a glance: queue status, delivery rates (1h/24h), connector health, alerts |
| **Queues** | Queue management with retry/suspend/remove, filtering, auto-refresh, and **recent errors panel** (FAIL/DEFER/DSN) |
| **Message Tracking** | Search `Get-MessageTrackingLog` with filters. **Color-coded EventId** (12 event types), **click to highlight all events with same MessageId** in blue. **Column chooser**, **CSV export**, Show Message Path, Cross-Server Trace |
| **Protocol Logs** | Parse SMTP Send/Receive protocol logs with **auto-detected paths from Exchange config**. Log type selector, recursive search |
| **Log Search** | Full-text regex search across any transport logs. **6 log types** auto-populated from Exchange config (Send/Receive Protocol, Message Tracking, Connectivity, Routing Table, Pipeline Tracing). UNC path support for remote servers |
| **Header Analyzer** | Parse email headers: Received hops with delay highlighting, SPF/DKIM/DMARC results, **ARC chain** (Arc-Seal/Arc-Authentication-Results), **multiple DKIM signatures**, TLS detection, X-Headers and notable headers |
| **Diagnostics** | DNS mail health (MX/SPF/DKIM/DMARC), Transport Rules viewer, Certificate manager, Connectivity Logs |
| **Statistics** | Mail flow analytics by Sender, Recipient, Domain, Hour, Connector. Sizes in **MB** |
| **Reports** | 9 report types: Full, Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Pipeline, BackPressure |

### Key Capabilities

- **EMS Auto-Detection** - detects Exchange Management Shell, loads snap-in into runspaces, auto-discovers local server
- **Server Scope** - auto-discovers all Exchange transport servers on connect. Query a specific server or all at once
- **Log Path Auto-Detection** - queries `Get-TransportService` for configured log locations, converts to UNC for remote access
- **Async Architecture** - all Exchange operations run in PowerShell runspaces with progress bars and a collapsible job console
- **Export** - all tabs support export to CSV/JSON. Message Tracking has dedicated CSV export with column filtering

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
|  ARC, DKIM)     |
+-----------------+

+-----------------+
| DNS Records     |     Tab: Diagnostics > DNS
| MX/SPF/DKIM/   |     Domain mail health check
| DMARC           |
+-----------------+

+-----------------+
| Transport Logs  |     Tab: Log Search
| (connectivity,  |     Regex search + context
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
When launched inside **Exchange Management Shell**, the tool:
1. Auto-detects EMS and loads the Exchange snap-in into async runspaces
2. Auto-discovers the local Exchange server name
3. Auto-connects and populates all log paths from transport config
4. Navigate tabs to monitor and manage transport
5. Watch the **Job Console** at the bottom for operation progress

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

# Export to CSV
.\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01 -OutputFile results.csv -OutputFormat CSV
```

### Environment Variable
```powershell
$env:EXCHANGE_SERVER = "exchange01.domain.local"
.\ExchangeRetry.ps1  # auto-fills server field
```

## Message Tracking EventId Colors

| EventId | Color | Meaning |
|---------|-------|---------|
| DELIVER | Green | Message delivered to mailbox |
| SEND | Light blue | Message sent to next hop |
| RECEIVE | Lavender | Message received from source |
| SUBMIT | Pale blue | Message submitted to pipeline |
| FAIL | Red | Delivery failed |
| DEFER | Yellow | Delivery deferred |
| DSN | Orange | Delivery status notification |
| RESOLVE | Mint | Recipient resolved |
| EXPAND | Lilac | Distribution group expanded |
| REDIRECT | Pink | Message redirected |
| TRANSFER | Teal | Message transferred |
| POISONMESSAGE | Dark red | Poison message detected |

Click any row to **highlight all events with the same MessageId** in blue bold.

## Async Architecture

```
+-------------------------------------------------+
| WinForms UI Thread                               |
|  +----------+  +----------+  +---------+        |
|  | Button   |  | Timer    |  | Poller  | (200ms)|
|  | Click    |  | Auto-    |  | checks  |        |
|  |          |  | refresh  |  | jobs    |        |
|  +----+-----+  +----+-----+  +----+----+        |
|       |              |            |              |
|       v              v            v              |
|  Start-AsyncJob    Start-     Update-AsyncJobs   |
|  (creates          AsyncJob   (EndInvoke,        |
|   runspace +                   callbacks)         |
|   EMS snap-in)                                    |
+---------+------------+-----------+---------------+
          |            |           |
          v            v           v
  +------------+  +------------+  +-----------------+
  | Runspace   |  | Runspace   |  | Job Console     |
  | (Exchange  |  | (Exchange  |  | [12:30] START #1|
  |  snap-in + |  |  snap-in + |  | [12:31] DONE #1|
  |  lib/*.ps1)|  |  lib/*.ps1)|  +-----------------+
  +------------+  +------------+
```

- Every Exchange operation runs in its own runspace with the Exchange snap-in loaded
- UI thread stays responsive - never blocks on network calls
- Progress bar shows when jobs are running
- Auto-refresh skips if jobs are already running (prevents pileup)
- All callbacks wrapped in try/catch - core never crashes from module errors

## Project Structure

```
ExchangeRetry/
+-- ExchangeRetry.ps1           # GUI (WinForms) - 9 tabs, async, job console
+-- ExchangeTrace.ps1           # CLI - same functions, console output
+-- lib/
|   +-- Core.ps1                # Exchange functions (connection, queues, tracking,
|   |                           #   headers, log parsing, reports, log path detection)
|   +-- Diagnostics.ps1         # DNS, transport rules, certificates, connectivity logs
|   +-- Monitoring.ps1          # Settings, cache, operator log, alerts, statistics
|   +-- AsyncRunner.ps1         # Async framework (runspaces, EMS snap-in, progress,
|                               #   job console, poller timer)
+-- tests/
|   +-- ExchangeRetry.Tests.ps1 # Pester tests for GUI functions
|   +-- ExchangeTrace.Tests.ps1 # Pester tests for CLI functions
+-- docs/
|   +-- architecture.md         # GUI layout diagrams, tab structure
|   +-- data-flow.md            # Data flow diagrams for all features
+-- CLAUDE.md                   # Dev instructions
+-- .env.example                # Environment variables template
+-- .gitignore
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

**0.6.0** - EMS auto-detection with snap-in loading in runspaces. Auto-discover Exchange server and transport log paths. Message Tracking: color-coded EventId (12 types), MessageId click-highlight, column chooser, CSV export. Header Analyzer: ARC chain, multiple DKIM signatures, notable headers. Log parsing with recursive search. Statistics in MB.

**0.5.0** - Full async architecture with runspaces, job console, progress bars. 9 tabs: Dashboard, Queues+Errors, Message Tracking with Path/Cross-Server, Protocol Logs, Log Search, Header Analyzer, Diagnostics (DNS/Rules/Certs/ConnLogs), Statistics, Reports.
