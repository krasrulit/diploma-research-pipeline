# Processed Data

Сюда pipeline сохраняет финальные Excel-файлы и промежуточные CSV.

## Основные файлы

- `additional_public_data.xlsx` — итоговый workbook по открытым источникам.
- `additional_public_data_manual_review.xlsx` — чеклист спорных сопоставлений для внешней ручной проверки.
- `additional_public_data_manual_review_validated.xlsx` — результат внешней ручной проверки, который не генерируется основным pipeline.
- `additional_public_data_validated.xlsx` — workbook после применения ручной проверки к MOEX / T-Invest слоям.
- `spark_combined_report.xlsx` — итоговый workbook по DOCX-отчетам СПАРК.
- `companies_master.csv` — нормализованный shortlist.
- `moex_instruments.csv` — все найденные кандидаты из MOEX.
- `moex_bonds.csv` — выбранные облигации MOEX.
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
