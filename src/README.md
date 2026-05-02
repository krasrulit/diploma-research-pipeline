# Source Code

В этой папке лежит весь рабочий код проекта.

## Модули публичного pipeline

- `load_shortlist.py`
  Читает shortlist Excel, автоматически ищет релевантные листы, нормализует `company_name`, `inn`, `ticker`, `ogrn`, `isin` и собирает `companies_master`.

- `load_moex.py`
  Работает с официальным MOEX ISS API, ищет акции и облигации, строит `moex_instruments`, `moex_bonds`, опциональные `moex_bond_history` / `moex_share_history` и лог сопоставления.

- `load_tinvest.py`
  Работает с официальным T-Invest API, тянет `Shares` и `Bonds`, матчится с shortlist по `ticker / isin / name`, а при включенном флаге выгружает еще и купонный календарь.

- `load_cbr.py`
  Загружает ключевую ставку и инфляцию с сайта Банка России.

- `load_commodities.py`
  Находит актуальные Excel-файлы Pink Sheet на сайте World Bank и парсит месячные и годовые commodity prices.

- `load_cbonds.py`
  Опциональный источник. Работает только если заданы credentials через переменные окружения. Без них pipeline не падает.

- `merge_public_data.py`
  Главный оркестратор публичного pipeline. Собирает результаты всех loaders, сохраняет промежуточные CSV и финальный Excel.

- `build_manual_review_workbook.py`
  Собирает Excel-чеклист спорных MOEX / T-Invest сопоставлений для ручной проверки.

- `apply_manual_validation.py`
  Применяет внешний файл ручной валидации к `additional_public_data.xlsx` и собирает `additional_public_data_validated.xlsx`.

- `build_public_market_workbook.py`
  Собирает единый `public_market_data_full.xlsx` без СПАРК-отчетностей: shortlist, validated mapping, MOEX, T-Invest, CBR, commodities и compact flags по компаниям. В `macro_cbr` теперь также входит `USDRUB`.

- `build_public_securities_workbook.py`
  Собирает единый security-universe workbook по всем компаниям: нормализует MOEX + T-Invest инструменты, считает history coverage, строит `security_source_resolution`, отдельный `security_resolution_clean`, company-level `company_market_access_clean`, применяет ручную проверку акций из `share_manual_review_quotes.xlsx`, подтягивает MOEX-review котировки из `moex_quotes_review/`, выделяет `security_manual_review` и сохраняет sidecar CSV по большим history-таблицам.

- `build_analysis_panel.py`
  Собирает объединенный SPARK-workbook по нефтегазу и металлургии, добавляет `sector` и `sample_flag`, подтягивает квартальные макро/commodity признаки, включая квартальные агрегаты `USDRUB`, добавляет Cbonds-календарь событий по облигациям и формирует итоговый `analysis_panel.xlsx`.

- `build_cbonds_event_calendar.py`
  Парсит вручную выгруженный `Календарь_событий.xlsx` из Cbonds, маппит события по ISIN и имени эмитента на компании из `companies_master`, строит квартальные event-признаки и файл контроля компаний, у которых есть Cbonds-облигации, но нет календаря событий.

- `cbonds_company_cards_capture.py`
  Вспомогательный macOS/Excel automation-скрипт для выгрузки карточек компаний из Cbonds add-in. Используется только локально, требует Microsoft Excel, установленный Cbonds add-in и разрешения Accessibility.

- `build_model_ready_panels.py`
  Берет `analysis_panel.xlsx` и формирует финальные датасеты для эконометрики: `model_ready_quarterly.xlsx`, `model_ready_annual.xlsx` и сводный `regression_ready_final.xlsx`.

- `run_h1_regression.py`
  Проверяет первую базовую гипотезу диплома на `regression_ready_final.xlsx`: готовит missingness/descriptive diagnostics и оценивает pooled OLS / year FE / firm FE / firm + year FE с clustered standard errors на уровне компании.

- `build_ownership_state_table.py`
  Собирает `ownership_state_table.xlsx` из raw СПАРК DOCX: вытягивает `Головная компания`, число `Дочерние компании`, строит `ownership_role`, inferred group и эвристический `state_owned_flag` с confidence/source-полями.

- `utils.py`
  Общие функции: нормализация текста, HTTP, логирование, сохранение CSV, загрузка локального `.env`.

## Модуль СПАРК

- `spark_docx_parser.py`
  Парсер `.docx`-отчетов СПАРК. Вытаскивает баланс, ОФР и ОДДС, строит `input_inventory`, `raw_long`, `panel_core`, `panel_quarterly`, `coverage`, `parse_log`. Для квартальной панели `FY` нормализуется в `Q4`, а по отсутствующим кварталам остаются пропуски.

- `build_spark_recovery_audit.py`
  Строит отдельный контрольный workbook по уже готовым `raw_long`: показывает core-метрики, которые можно восстановить без `metric_code`, например `debt_lt` / `debt_st` по связке `Заёмные средства` + раздел баланса `IV/V обязательства`.

- `build_spark_regression_patch.py`
  Берет `spark_metric_recovery_audit.xlsx` и `regression_ready_final.xlsx`, заполняет все восстановленные SPARK-показатели в формате финальной панели, пересчитывает зависимые коэффициенты и сохраняет `spark_regression_ready_patch.xlsx` с same-column листами для ручной вставки.

## Точки входа

- [main_public_data.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_data.py)
- [main_public_market.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_market.py)
- [main_public_securities.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_securities.py)
- [main_ownership_state.py](/Users/grigorijkrasovickij/Documents/Playground/main_ownership_state.py)
- [main_spark.py](/Users/grigorijkrasovickij/Documents/Playground/main_spark.py)
- [main_spark_recovery_audit.py](/Users/grigorijkrasovickij/Documents/Playground/main_spark_recovery_audit.py)
- [main_spark_regression_patch.py](/Users/grigorijkrasovickij/Documents/Playground/main_spark_regression_patch.py)
- [main_analysis_panel.py](/Users/grigorijkrasovickij/Documents/Playground/main_analysis_panel.py)
- [main_cbonds_event_calendar.py](/Users/grigorijkrasovickij/Documents/Playground/main_cbonds_event_calendar.py)
- [main_cbonds_company_cards_capture.py](/Users/grigorijkrasovickij/Documents/Playground/main_cbonds_company_cards_capture.py)
- [main_model_ready.py](/Users/grigorijkrasovickij/Documents/Playground/main_model_ready.py)
- [main_h1_regression.py](/Users/grigorijkrasovickij/Documents/Playground/main_h1_regression.py)
