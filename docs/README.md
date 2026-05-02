# Project Documentation

Эта папка содержит человекочитаемые описания пайплайнов, которые дополняют код из корня репозитория и `src/`.

## Документы

- `SPARK_PIPELINE.md` — как устроен парсер СПАРК, какие входные файлы нужны, какие листы он формирует и как проверять результат.

## Где лежит код

- Основные entrypoint-файлы лежат в корне репозитория: `main_*.py`.
- Рабочие модули лежат в `src/`.
- Парсер СПАРК: `src/spark_docx_parser.py`.
- Сборка финальной панели: `src/build_analysis_panel.py`.
- Сборка model-ready файлов: `src/build_model_ready_panels.py`.
- Сборка public securities слоя: `src/build_public_securities_workbook.py`.
