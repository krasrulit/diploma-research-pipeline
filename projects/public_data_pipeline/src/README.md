# Source Code

Эта папка содержит весь исполняемый код pipeline.

## Файлы

- `load_shortlist.py` — читает shortlist Excel, нормализует идентификаторы и строит `companies_master`.
- `load_moex.py` — ищет и сопоставляет инструменты через MOEX ISS, вытягивает облигации и лог сопоставления.
- `load_cbr.py` — загружает ключевую ставку и инфляцию с официального сайта Банка России.
- `load_commodities.py` — скачивает и парсит Pink Sheet файлы World Bank.
- `load_cbonds.py` — optional интеграция с Cbonds, не ломается без credentials.
- `merge_all.py` — оркестрация pipeline и сборка итогового Excel.
- `utils.py` — общие утилиты: нормализация, HTTP, логирование, сохранение таблиц.
