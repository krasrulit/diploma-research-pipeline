# Processed Data

Сюда pipeline сохраняет финальные Excel-файлы и промежуточные CSV.

## Основные файлы

- `additional_public_data.xlsx` — итоговый workbook по открытым источникам.
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
