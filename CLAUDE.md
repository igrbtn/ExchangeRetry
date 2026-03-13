# ExchangeRetry

## Overview
Полнофункциональный GUI + CLI для мониторинга и управления транспортом Microsoft Exchange. Отслеживание сообщений на каждом этапе пайплайна: очереди, message tracking, SMTP protocol logs, transport logs, анализ заголовков, отчёты. Все операции выполняются асинхронно через runspaces — GUI никогда не зависает.

## Quick Start
```powershell
# GUI (полный функционал)
.\ExchangeRetry.ps1

# CLI — парсинг заголовков
.\ExchangeTrace.ps1 -HeaderFile .\header.txt

# CLI — трассировка по Message-ID
.\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01

# CLI — поиск по транспортным логам
.\ExchangeTrace.ps1 -TransportLogPath "\\server\TransportLogs" -SearchPattern "user@domain.com"

# CLI — отчёт
.\ExchangeTrace.ps1 -Report Full -Server exchange01

# Тесты
Invoke-Pester -Path ./tests/
```

## Architecture

### Модульная структура (lib/):
- `Core.ps1` — Exchange-функции: подключение, очереди, message tracking, headers, reports
- `Diagnostics.ps1` — DNS (MX/SPF/DKIM/DMARC), transport rules, certificates, connectivity logs
- `Monitoring.ps1` — settings, cache, operator log, alerts, statistics
- `AsyncRunner.ps1` — async framework: runspaces, job tracker, progress bar, job console

### GUI (ExchangeRetry.ps1) — 9 вкладок:
1. **Dashboard** — здоровье транспорта: очереди, delivery rate (1h/24h), коннекторы, алерты
2. **Queues** — управление очередями: retry/suspend/remove с фильтрацией, auto-refresh, панель ошибок (FAIL/DEFER/DSN)
3. **Message Tracking** — Get-MessageTrackingLog с фильтрами + Show Message Path + Cross-Server Trace
4. **Protocol Logs** — парсинг SMTP Send/Receive protocol logs
5. **Log Search** — текстовый поиск по любым логам с контекстом
6. **Header Analyzer** — парсинг email-заголовков: hops, delays, SPF/DKIM/DMARC, TLS, X-Headers
7. **Diagnostics** — DNS health check, Transport Rules, Certificates, Connectivity Logs
8. **Statistics** — по отправителям, получателям, доменам, часам, коннекторам
9. **Reports** — Full, Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Pipeline, BackPressure

### Async Architecture:
- Все Exchange-операции выполняются в PowerShell runspaces (Start-AsyncJob)
- WinForms Timer (200ms) опрашивает завершение jobs и вызывает UI-callbacks
- Прогресс-бар и счетчик running jobs в нижней панели
- Collapsible job console с логом всех операций
- Comprehensive try/catch — ядро не падает при ошибках модулей/связи

### CLI (ExchangeTrace.ps1):
- Те же функции в консольном режиме с цветным выводом
- Экспорт в CSV/JSON

## Configuration
- `EXCHANGE_SERVER` — FQDN Exchange-сервера (env или ввод в GUI)
- Auto-refresh: 30 секунд (настраивается). Не запускается если есть running jobs
- Settings persist в AppData (JSON)

## Testing
- Framework: Pester v5
- Тесты: `tests/ExchangeRetry.Tests.ps1`, `tests/ExchangeTrace.Tests.ps1`

## Versioning
- Текущая версия: 0.5.0
