# ExchangeRetry

## Overview
Полнофункциональный GUI + CLI для мониторинга и управления транспортом Microsoft Exchange. Отслеживание сообщений на каждом этапе пайплайна: очереди, message tracking, SMTP protocol logs, transport logs, анализ заголовков, отчёты.

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

### GUI (ExchangeRetry.ps1) — 7 вкладок:
1. **Dashboard** — здоровье транспорта: очереди, delivery rate (1h/24h), коннекторы, ошибки
2. **Queues** — управление очередями: retry/suspend/remove с фильтрацией и auto-refresh
3. **Message Tracking** — Get-MessageTrackingLog с фильтрами: EventId, Sender, Recipient, Subject, даты. Кнопка "Show Message Path" — визуализация маршрута письма
4. **Protocol Logs** — парсинг SMTP Send/Receive protocol logs (CSV с #-комментариями)
5. **Log Search** — текстовый поиск по любым логам с контекстом
6. **Header Analyzer** — парсинг email-заголовков: hops, delays, SPF/DKIM/DMARC, TLS, X-Headers
7. **Reports** — Full, Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Pipeline, BackPressure

### CLI (ExchangeTrace.ps1):
- Те же функции в консольном режиме с цветным выводом
- Экспорт в CSV/JSON

## Configuration
- `EXCHANGE_SERVER` — FQDN Exchange-сервера (env или ввод в GUI)
- Auto-refresh: 30 секунд (настраивается в `$script:Config`)

## Testing
- Framework: Pester v5
- Тесты: `tests/ExchangeRetry.Tests.ps1`, `tests/ExchangeTrace.Tests.ps1`

## Versioning
- Текущая версия: 0.3.0
