# Processed Data

Сюда pipeline сохраняет финальные Excel-файлы и промежуточные CSV.

## Основные файлы

- `additional_public_data.xlsx` — итоговый workbook по открытым источникам.
- `additional_public_data_manual_review.xlsx` — чеклист спорных сопоставлений для внешней ручной проверки.
- `additional_public_data_manual_review_validated.xlsx` — результат внешней ручной проверки, который не генерируется основным pipeline.
- `additional_public_data_validated.xlsx` — workbook после применения ручной проверки к MOEX / T-Invest слоям.
- `public_market_data_full.xlsx` — полный workbook по публичным источникам без СПАРК-отчетностей.
- `spark_combined_report.xlsx` — итоговый workbook по DOCX-отчетам СПАРК.
- `spark_combined_report_2014q3_2025q4.xlsx` — вариант СПАРК workbook с квартальной сеткой `2014Q3–2025Q4` по всем компаниям из входной папки.
- `spark_neftegaz_report_2014q3_2025q4.xlsx` — нефтегазовый SPARK workbook.
- `spark_metallurgy_report_2014q3_2025q4.xlsx` — металлургический SPARK workbook.
- `spark_sector_combined_2014q3_2025q4.xlsx` — объединенный SPARK workbook по обоим секторам с признаками `sector` и `sample_flag`.
- `analysis_panel.xlsx` — итоговая исследовательская квартальная панель.
- `companies_master.csv` — нормализованный shortlist.
- `moex_instruments.csv` — все найденные кандидаты из MOEX.
- `moex_bonds.csv` — выбранные облигации MOEX.
- `moex_bond_history.csv` — история торгов/доходностей MOEX по выбранным облигациям, если включена.
- `moex_share_history.csv` — история торгов MOEX по выбранным акциям, если включена.
- `tinvest_instruments.csv` — все найденные кандидаты из T-Invest.
- `tinvest_bonds.csv` — выбранные облигации T-Invest.
- `tinvest_bond_coupons.csv` — купонный календарь из T-Invest, если включен флаг.
- `macro_cbr.csv` — ряды Банка России.
- `commodity_prices_monthly.csv` — месячные commodity prices.
- `commodity_prices_annual.csv` — годовые commodity prices.
- `mapping_log.csv` — лог сопоставлений.
- `download_log.csv` — лог скачивания и ошибок.

Если какого-то файла нет, это обычно значит, что соответствующий источник не запускался или был отключен.

## Как пользоваться validated workbook

Начинайте с листа `validation_summary`, затем переходите в `issuer_mapping_validated`. Именно этот лист удобнее использовать как очищенный слой связей `company -> instrument` после ручной проверки.

Если нужно пересобрать файл после новой ручной проверки:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 src/apply_manual_validation.py \
  --main-workbook data_processed/additional_public_data.xlsx \
  --validation-workbook data_processed/additional_public_data_manual_review_validated.xlsx \
  --output data_processed/additional_public_data_validated.xlsx
```

Для работы без отчетностей удобнее открывать `public_market_data_full.xlsx`: там есть `firm_market_flags` и все публичные слои в одном месте.

Для отчетностей по СПАРК удобнее открывать `spark_combined_report_2014q3_2025q4.xlsx`: там есть `panel_quarterly`, где каждая компания разложена по всем кварталам диапазона, а пропуски оставлены пустыми.

Для межотраслевого анализа удобнее открывать `analysis_panel.xlsx`: там уже объединены нефтегаз и металлургия, добавлены `sector`, `sample_flag`, квартальные макро- и commodity-признаки, а для нефтегаза еще и public-market flags из открытых источников.
