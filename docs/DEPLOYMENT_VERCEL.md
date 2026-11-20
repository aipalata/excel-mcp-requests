# Подготовка к Vercel

Локальная версия использует stdio и книгу на компьютере. В облаке исходная книга хранится в приватном бакете Supabase Storage по пути `workbooks/requests.xlsx`. Python-функция Vercel скачивает её во временный каталог для вызова инструмента и запускает готовый Excel MCP Server 2.3.2 со Streamable HTTP в режиме чтения. Между запросами файл хранится в Supabase. [Файлы в Vercel Functions](https://vercel.com/kb/guide/how-can-i-use-files-in-serverless-functions) · [Приватные бакеты](https://supabase.com/docs/guides/storage/buckets/fundamentals)

## Чтение в облаке

- Настроить `MCP_MODE=http`, `MCP_HTTP_URL=https://excel-mcp-requests.vercel.app/api/workbook_mcp` и `EXCEL_CLOUD_READ_ONLY=1` в production-проекте Vercel.
- Включить Vercel Authentication для сгенерированного адреса развёртывания. На этом адресе таблица вызывает `initialize`, `tools/list`, `describe_sheet` и `read_range`. Канонический публичный адрес показывает интерфейс.
- В Python-функции задать идентификатор проекта и publishable key через `WORKBOOK_SUPABASE_PROJECT` и `WORKBOOK_SUPABASE_PUBLISHABLE_KEY`. Секрет `WORKBOOK_STORAGE_TOKEN` передаётся Supabase только по HTTPS. Политика Storage сравнивает его хеш и разрешает лишь SELECT для `workbooks/requests.xlsx`. Для доступа к самому MCP используется отдельный серверный `MCP_HTTP_TOKEN`.

Проверить MCP-адаптер локально можно командой `.venv/bin/python -m unittest discover -s tests -p 'workbook_mcp_test.py'`. Тест использует готовый сервер и подставляет локальную книгу на месте скачивания из Storage.

## Подготовка записи и облачного чата

1. Добавить вход для сотрудников и ограничение частоты запросов.
2. Перенести планы, историю и идемпотентность в постоянное хранилище. Обновлять объект книги с контролем версии, чтобы одна запись не перезаписывала другую.
3. Выдать отдельный серверный OpenAI API Key для облака и предоставить облачному MCP узкий доступ на обновление книги.
4. Проверить все три сценария, отмену, устаревший план, повтор запроса, параллельные подтверждения и сохранение при новом запуске.

Пока Vercel принимает только запросы обновления таблицы. Все операции с записью по-прежнему доступны в локальном режиме.
