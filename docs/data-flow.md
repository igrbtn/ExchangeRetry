# Data Flow

## Connection & Server Discovery

```
User enters server FQDN
        |
        v
Connect-ExchangeRemote
  New-PSSession (Kerberos)
  Import-PSSession
        |
        v
Get-TransportService
  Discover all transport servers
        |
        v
Populate Scope combo
  (All Servers) | EX01 | EX02 | ...
```

## Queue Management Flow

```
                    Scope: [Server / All]
                           |
              +------------+------------+
              |                         |
              v                         v
        Get-Queue                 Get-Queue
        -Server EX01              (no filter)
              |                         |
              +------------+------------+
                           |
                           v
                    DataGridView
                    (queue list)
                           |
                    User selects queue
                           |
                           v
                    Get-Message
                    -Queue <identity>
                           |
                           v
                    DataGridView
                    (message list)
                           |
              +------+-----+------+
              |      |            |
              v      v            v
         Retry-   Suspend-    Remove-
         Queue    Message     Message
              |      |            |
              v      v            v
              Refresh all grids
```

## Error Tracking Flow

```
Get-MessageTrackingLog
  EventId = FAIL     ─┐
  EventId = DEFER    ─┼─> Merge + Sort by Timestamp DESC
  EventId = DSN      ─┘
        |
        v
  DataGridView (errors panel)
    Time | Event | Sender | Recipients | Subject | Server | Error
        |
        v
  Summary: FAIL: N | DEFER: N | DSN: N | Total: N
```

## Message Tracking Flow

```
User input:
  MessageId / Sender / Recipient / Subject
  EventId filter / Date range
        |
        v
Get-MessageTrackingLog
  (with all filters)
        |
        v
DataGridView (tracking results)
        |
        +--------> [Show Message Path]
        |               |
        |               v
        |          Filter by selected MessageId
        |          Sort by Timestamp ASC
        |               |
        |               v
        |          Popup window with step-by-step path:
        |            [1] RECEIVE -> [2] SUBMIT -> [3] SEND -> [4] DELIVER
        |            with connectors, IPs, contexts
        |
        +--------> [Export]
                        |
                        v
                   CSV / JSON file
```

## Protocol Log Parsing Flow

```
User selects log directory
  (SendProtocolLog / ReceiveProtocolLog)
        |
        v
Get-ChildItem *.LOG
  (newest N files)
        |
        v
For each file:
  Parse #Fields: header -> column names
  Parse CSV data lines
  Apply text filter across all columns
        |
        v
DataGridView (parsed entries)
  date-time | connector-id | session-id | sequence-number | ...
  local-endpoint | remote-endpoint | event | data | context
```

## Header Analysis Flow

```
User pastes headers / loads file
        |
        v
Parse-EmailHeaders
  |
  +-- Extract: Message-ID, From, To, Subject, Date, Return-Path
  +-- Extract: SPF, DKIM, DMARC results
  +-- Extract: X-MS-Exchange-* headers
  +-- Parse Received: headers
  |     Unfold multiline -> extract From/By/With/For/TLS/Timestamp
  |     Reverse to chronological order
  |     Calculate delays between hops
  |
  v
Display:
  Info panel (Message-ID, From/To, auth results, total delay)
  Hops grid (# | From | By | Protocol | TLS | Timestamp | Delay)
  X-Headers grid (Header | Value)
```

## Transport Reports Flow

```
User selects report type
        |
        v
Get-TransportReportData
  |
  +-- Queues:      Get-Queue -> summary by status
  +-- Connectors:  Get-SendConnector + Get-ReceiveConnector
  +-- AgentLog:    Get-TransportAgent + Get-AgentLog
  +-- RoutingTable: Get-TransportConfig + Get-TransportService
  +-- DSN:         Get-TransportConfig + Get-MessageTrackingLog (DSN)
  +-- Summary:     Get-MessageTrackingLog -> stats, top senders, failures
  +-- Pipeline:    Get-TransportAgent (sorted by priority)
  +-- BackPressure: Get-ExchangeDiagnosticInfo (ResourceThrottling)
  +-- Full:        All of the above
        |
        v
Console-style text output in dark theme textbox
        |
        v
[Save to File] -> .txt / .json
```
