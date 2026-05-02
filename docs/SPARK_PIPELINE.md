# SPARK DOCX Pipeline

Этот документ описывает слой, который превращает `.docx`-отчеты СПАРК в квартальную панель для дипломной модели.

## Главные файлы

- `main_spark.py` — короткая точка входа, вызывает `src.spark_docx_parser.main`.
- `src/spark_docx_parser.py` — полный парсер СПАРК DOCX.
- `main_spark_recovery_audit.py` — собирает audit workbook по строкам, которые можно восстановить без `metric_code`.
- `main_spark_regression_patch.py` — превращает восстановленные строки в готовый файл-вставку для `regression_ready_final.xlsx`.
- `main_analysis_panel.py` — объединяет готовые SPARK workbook по секторам с макро, commodities, ownership, Cbonds и public securities.
- `main_model_ready.py` — превращает `analysis_panel.xlsx` в квартальную и годовую model-ready панели.

## Какие входные файлы ожидает парсер

Парсер ищет файлы вида:

```text
СПАРК-Отчет_<название компании>_<ИНН>_<дата>_<номер>.docx
```

Поддерживаемый формат сейчас только `.docx`. PDF-файлы не парсятся: они попадают в инвентарь/лог как неподдержанные, чтобы было видно, что файл лежал в папке, но не был использован.

## Какие формы парсятся

Парсер вытаскивает таблицы:

- `Бухгалтерский баланс`;
- `Отчет о финансовых результатах`;
- `Отчет о движении денежных средств`.

Ключевые строки определяются по кодам РСБУ и названиям строк. Например:

- `1600` -> `assets_total`;
- `1150` -> `ppe`;
- `1250` -> `cash`;
- `1300` -> `equity`;
- `1410` -> `debt_lt`;
- `1510` -> `debt_st`;
- `2110` -> `revenue`;
- `2330` -> `interest_expense`;
- `2400` -> `net_income`;
- `4100` -> `cfo`;
- `4221` -> `capex`;
- `4311` -> `loans_received`;
- `4314` -> `bonds_issued`;
- `4323` -> `debt_repayment`.

## Как запускать отдельный SPARK workbook

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_spark.py \
  --input-dir "/path/to/spark/docx/folder" \
  --pattern "СПАРК-Отчет_*.docx" \
  --quarter-grid-start 2014Q3 \
  --quarter-grid-end 2025Q4 \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_combined_report.xlsx"
```

## Какие листы формирует SPARK workbook

- `input_inventory` — список найденных файлов, исходный путь, расширение, компания, ИНН, флаг поддерживаемого формата.
- `raw_long` — все распарсенные строки отчетности в длинном формате.
- `panel_core` — одна строка на компанию/период с ключевыми финансовыми показателями.
- `panel_quarterly` — регулярная квартальная сетка по всем компаниям и кварталам диапазона; пропуски остаются пустыми.
- `coverage` — покрытие по компаниям, периодам и формам отчетности.
- `parse_log` — технический лог: какие файлы распарсились, какие были пропущены и почему.

## Как годовые периоды превращаются в кварталы

Если в СПАРК период указан как `FY`, парсер нормализует его в `Q4`. Поэтому годовая панель дальше строится как Q4-срез квартальной панели.

## Как проверять результат

Сначала смотрим:

- `parse_log` — нет ли массовых ошибок парсинга;
- `coverage` — какие компании и периоды реально покрыты;
- `panel_quarterly` — нет ли неожиданных дублей `company_id + quarter_label`;
- `input_inventory` — не лежат ли в папке PDF или лишние файлы, которые ожидаемо не попали в данные.

## Восстановление строк без `metric_code`

В некоторых отчетах СПАРК код строки не указан, но смысл строки можно восстановить по названию строки и разделу баланса. Например, если:

- `table_title = Бухгалтерский баланс`;
- `statement_section = IV. ДОЛГОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА`;
- `metric_name = Заёмные средства`;
- `metric_code` пустой,

то строка классифицируется как `debt_lt`.

Аналогично:

- `statement_section = V. КРАТКОСРОЧНЫЕ ОБЯЗАТЕЛЬСТВА`;
- `metric_name = Заёмные средства`;

классифицируется как `debt_st`.

Если не нужно пересобирать весь SPARK workbook, можно построить отдельный audit-файл:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_spark_recovery_audit.py \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_metric_recovery_audit.xlsx"
```

В нем:

- `recovery_needed_long` — только значения, которых раньше не было в `panel_core`, но которые восстановлены новыми fallback-правилами;
- `recovery_needed_wide` — тот же блок в широком формате для ручной вставки;
- `debt_recovery_long` / `debt_recovery_wide` — отдельный блок только по `debt_lt` и `debt_st`;
- `recovered_raw_values` — все raw-строки, распознанные по fallback-правилам;
- `panel_before_after` — сравнение старого и нового core-слоя.

Чтобы получить не отдельный debt-блок, а готовые строки для финального regression workbook по всем восстановленным показателям:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_spark_regression_patch.py \
  --regression-ready "/Users/grigorijkrasovickij/Documents/Playground/data_processed/regression_ready_final.xlsx" \
  --recovery-audit "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_metric_recovery_audit.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_regression_ready_patch.xlsx"
```

Этот файл не перезаписывает исходную модель. Он делает безопасный patch-layer:

- заполняет только пустые финансовые значения из СПАРК;
- не трогает уже заполненные значения;
- пересчитывает `total_debt`, debt ratios, `net_debt`, `net_debt_to_assets`, `available_core_metrics_count`, `reporting_observed_flag` и другие зависимые поля;
- сохраняет строки в листах `quarterly_patch_same_columns` и `annual_patch_same_columns` с теми же названиями и порядком колонок, что в `regression_ready_final.xlsx`.

Для финальной модели основной файл проверки:

```text
data_processed/regression_ready_final.xlsx
```

В нем лист `variable_dictionary` показывает, что лежит в каждой колонке, источник, исходный файл/URL, лист и способ расчета.

## Важные ограничения

- Парсер не делает ручное заполнение пропусков.
- Парсер не интерпретирует PDF.
- Парсер оставляет пропуски как пропуски; решение об импутации принимается уже на этапе эконометрической модели.
- Если формат DOCX сильно отличается от стандартного СПАРК-отчета, строка может уйти в `parse_log` как проблемная и потребовать ручной проверки.
