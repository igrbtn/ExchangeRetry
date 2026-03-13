# Architecture

## Transport Pipeline Coverage

```
                        ExchangeRetry Coverage
                        ======================

  Internet/Internal
        |
        v
  +------------------+     +------------------+
  | Receive          |     | Protocol Logs    |  <-- Tab: Protocol Logs
  | Connector        |---->| (SMTP Receive)   |      Parse & filter
  | (SMTP/TLS)       |     +------------------+
  +------------------+
        |
        v
  +------------------+     +------------------+
  | Transport        |     | Agent Log        |  <-- Tab: Reports > AgentLog
  | Agents           |---->| (actions/events) |      View agent decisions
  | (anti-spam, etc) |     +------------------+
  +------------------+
        |
        v
  +------------------+     +------------------+
  | Categorizer      |     | Message Tracking |  <-- Tab: Message Tracking
  | (routing,        |---->| RECEIVE/SUBMIT   |      Filter by EventId,
  |  resolution)     |     | RESOLVE/EXPAND   |      Sender, Recipient,
  +------------------+     +------------------+      Subject, dates
        |
        v
  +------------------+     +------------------+
  | Queue            |     | Queue Viewer     |  <-- Tab: Queues
  | (submission,     |---->| retry/suspend/   |      Manage messages,
  |  delivery,       |     | remove + errors  |      view FAIL/DEFER/DSN
  |  poison, shadow) |     +------------------+
  +------------------+
        |
        v
  +------------------+     +------------------+
  | Send             |     | Protocol Logs    |  <-- Tab: Protocol Logs
  | Connector        |---->| (SMTP Send)      |      Parse & filter
  | (SMTP/TLS)       |     +------------------+
  +------------------+
        |
        v
  +------------------+     +------------------+
  | Delivery /       |     | Message Tracking |  <-- Tab: Message Tracking
  | Relay            |---->| DELIVER/SEND/    |      "Show Message Path"
  |                  |     | FAIL/DSN/DEFER   |      full route view
  +------------------+     +------------------+

  +------------------+
  | Email Headers    |     <-- Tab: Header Analyzer
  | (Received hops,  |         Parse hops, delays,
  |  SPF/DKIM/DMARC, |         SPF/DKIM/DMARC,
  |  X-Headers, TLS) |         X-MS-Exchange-*
  +------------------+

  +------------------+
  | Transport Logs   |     <-- Tab: Log Search
  | (connectivity,   |         Text search with
  |  routing, any)   |         context display
  +------------------+
```

## GUI Tab Structure

```
+------------------------------------------------------------------+
| Exchange Server: [____________] [Connect]  Connected   Scope: [v] |
+------------------------------------------------------------------+
| Dashboard | Queues | Tracking | Protocol | Logs | Headers | Reports|
+------------------------------------------------------------------+
|                                                                    |
|  Tab content area                                                  |
|                                                                    |
+------------------------------------------------------------------+
| Status: Ready                                                      |
+------------------------------------------------------------------+
```

### Dashboard
```
+------------------------------------------------------------------+
| [Refresh Dashboard]                                                |
+------------------------------------------------------------------+
|                                                                    |
|   EXCHANGE TRANSPORT DASHBOARD                                     |
|   Server: exchange01                                               |
|   ============================================================     |
|                                                                    |
|   QUEUES                                                           |
|   Total queues: 12    Total messages: 45                           |
|     Ready: 8  Active: 3  Retry: 1 *** ATTENTION ***                |
|                                                                    |
|   DELIVERY (1h)                                                    |
|   Received: 234  Delivered: 228  Failed: 2  Deferred: 4           |
|   Delivery %: 99.1%                                               |
|                                                                    |
|   DELIVERY (24h)                                                   |
|   Received: 5420  Delivered: 5380  Failed: 15  Deferred: 25       |
|   Delivery %: 99.7%                                               |
|                                                                    |
|   CONNECTORS                                                       |
|   Send: 3 active, 1 disabled                                      |
|   Receive: 5 active, 0 disabled                                   |
|                                                                    |
+------------------------------------------------------------------+
```

### Queues (3-panel layout)
```
+------------------------------------------------------------------+
| Queues                                                             |
| [Refresh] [Retry Queue] [Filter___________] [x] Auto (30s)        |
+------------------------------------------------------------------+
| Identity     | Type    | Status | Count | NextHop    | LastError   |
|--------------|---------|--------|-------|------------|-------------|
| server\12    | SMTP    | Ready  | 5     | domain.com | -           |
| server\15    | SMTP    | Retry  | 12    | bad.com    | 451 timeout |
+------------------------------------------------------------------+
| Messages: server\15                                                |
| [Retry Selected] [Suspend Selected] [Remove Selected] [x]NDR      |
+------------------------------------------------------------------+
| Identity | From          | Status | Subject      | LastError       |
|----------|---------------|--------|--------------|-----------------|
| 15\1     | user@dom.com  | Retry  | Test email   | 451 timeout     |
+------------------------------------------------------------------+
| Recent Errors (FAIL/DEFER/DSN -- last 24h)                         |
| [Refresh Errors] [Export...]   FAIL: 3 | DEFER: 7 | DSN: 1        |
+------------------------------------------------------------------+
| Time                | Event | Sender      | Recipients | Error     |
|---------------------|-------|-------------|------------|-----------|
| 2026-03-13 14:22:01 | FAIL  | a@test.com  | b@bad.com  | 550 User  |
| 2026-03-13 14:20:15 | DEFER | c@test.com  | d@slow.com | 421 Try   |
+------------------------------------------------------------------+
```

### Message Tracking
```
+------------------------------------------------------------------+
| Message-ID: [________] Sender: [________] Recipient: [________]    |
| From: [2026-03-12 00:00] To: [2026-03-13 23:59] Event: [(All) v]  |
| Subject: [________]                                  [Search]      |
+------------------------------------------------------------------+
| Timestamp       | EventId  | Source | Sender     | Server | ...    |
|-----------------|----------|--------|------------|--------|--------|
| 14:22:01.123    | RECEIVE  | SMTP   | user@...   | EX01   | ...    |
| 14:22:01.456    | SUBMIT   | STORE  | user@...   | EX01   | ...    |
| 14:22:02.100    | DELIVER  | STORE  | user@...   | EX02   | ...    |
+------------------------------------------------------------------+
| [Export...] [Show Message Path]  3 event(s)                        |
+------------------------------------------------------------------+
```

### Message Path (popup)
```
+------------------------------------------------------------------+
| MESSAGE PATH: <abc123@domain.com>                                  |
| Subject: Test Message                                              |
| Sender: user@domain.com                                            |
|                                                                    |
|   [1] 14:22:01.123  RECEIVE      SMTP       EX01                  |
|       Connector : Default Frontend EX01                            |
|       Recipients: dest@target.com                                  |
|       ClientIP  : 203.0.113.10 (mail.sender.com)                  |
|                                                                    |
|   [2] 14:22:01.456  SUBMIT       STOREDRV   EX01                  |
|       Recipients: dest@target.com                                  |
|                                                                    |
|   [3] 14:22:02.100  SEND         SMTP       EX01                  |
|       Connector : Outbound to Internet                             |
|       Recipients: dest@target.com                                  |
|                                                                    |
|   [4] 14:22:03.200  DELIVER      STOREDRV   EX02                  |
|       Recipients: dest@target.com                                  |
+------------------------------------------------------------------+
```
