# ExchangeRetry

## Overview
Два инструмента для Microsoft Exchange: GUI для управления очередями (retry/suspend/remove) и консольная утилита для трассировки писем (парсинг заголовков, поиск по транспортным логам, отчёты).

## Quick Start
```powershell
# GUI для очередей
.\ExchangeRetry.ps1

# Трассировка: парсинг заголовков из файла
.\ExchangeTrace.ps1 -HeaderFile .\header.txt

# Трассировка: поиск по Message-ID
.\ExchangeTrace.ps1 -MessageId "<abc@domain.com>" -Server exchange01

# Трассировка: поиск по отправителю
.\ExchangeTrace.ps1 -Sender user@domain.com -Server exchange01

# Поиск по транспортным логам
.\ExchangeTrace.ps1 -TransportLogPath "\\exchange01\TransportLogs" -SearchPattern "user@domain.com"

# Полный отчёт по транспорту
.\ExchangeTrace.ps1 -Report Full -Server exchange01

# Тесты
Invoke-Pester -Path ./tests/
```

## Architecture
- **ExchangeRetry.ps1** — WinForms GUI: подключение к Exchange, просмотр очередей, retry/suspend/remove
- **ExchangeTrace.ps1** — CLI-утилита:
  - Парсинг email-заголовков (Received hops, SPF/DKIM/DMARC, X-Headers)
  - Message tracking через `Get-MessageTrackingLog`
  - Поиск по текстовым транспортным логам (SMTP Send/Receive)
  - Отчёты: Queues, Connectors, AgentLog, RoutingTable, DSN, Summary, Full
  - Экспорт в CSV/JSON

## Configuration
- `EXCHANGE_SERVER` — FQDN Exchange-сервера (env или ввод в GUI)
- Auto-refresh interval: 30 секунд (настраивается в `$script:Config`)

## Testing
- Framework: Pester v5
- Тесты: `tests/ExchangeRetry.Tests.ps1`
- Запуск: `Invoke-Pester`

## Versioning
- Текущая версия: 0.1.0 (в шапке скриптов `.NOTES`)
