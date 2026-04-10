# Diploma Research Pipeline

Репозиторий для двух практических задач диплома:

1. сборка общедоступных финансовых и рыночных данных по shortlist компаний;
2. парсинг `.docx`-выгрузок СПАРК в единую панель.

## Структура репозитория

- [main_public_data.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_data.py) — основной entrypoint для публичных источников.
- [main_public_market.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_market.py) — полный public-market workbook без СПАРК-отчетностей.
- [main_spark.py](/Users/grigorijkrasovickij/Documents/Playground/main_spark.py) — entrypoint для парсинга отчетов СПАРК.
- [src/README.md](/Users/grigorijkrasovickij/Documents/Playground/src/README.md) — подробная карта модулей и того, что делает каждый файл.
- [data_raw/README.md](/Users/grigorijkrasovickij/Documents/Playground/data_raw/README.md) — что хранить в сыром виде.
- [data_processed/README.md](/Users/grigorijkrasovickij/Documents/Playground/data_processed/README.md) — какие итоговые файлы появляются после запуска.
- [notebooks/README.md](/Users/grigorijkrasovickij/Documents/Playground/notebooks/README.md) — место для валидационных и исследовательских ноутбуков.
- [.env.example](/Users/grigorijkrasovickij/Documents/Playground/.env.example) — шаблон локальных секретов.
- [requirements.txt](/Users/grigorijkrasovickij/Documents/Playground/requirements.txt) — зависимости проекта.

## Что делать с токеном T-Invest

Токен нужно вставлять не в код и не в ноутбук, а в локальный файл `.env` в корне репозитория.

1. Скопируйте шаблон:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
cp .env.example .env
```

2. Откройте файл `.env`.

3. Вставьте токен в строку:

```bash
TINVEST_TOKEN=ваш_read_only_token
```

Файл `.env` уже исключен из git, поэтому токен не должен попасть в GitHub. `main_public_data.py` подхватывает `.env` автоматически при запуске.

Если T-Invest в вашей сети отвечает с ошибкой SSL-сертификата, добавьте в тот же `.env`:

```bash
TINVEST_SSL_VERIFY=false
```

Это нужно только для T-Invest и только если у вас локально есть SSL interception / proxy. Более аккуратный вариант: указать путь к вашему корпоративному CA bundle в `TINVEST_CA_BUNDLE`.

## Установка

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
pip install -r requirements.txt
```

## Запуск 1. Публичные данные

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
python main_public_data.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/additional_public_data.xlsx" \
  --with-tinvest
```

Если хотите еще и купонный календарь по облигациям из T-Invest:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
python main_public_data.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/additional_public_data.xlsx" \
  --with-tinvest \
  --tinvest-coupons
```

## Запуск 1B. Полный public-market файл без отчетностей

Этот режим собирает публичные источники, T-Invest, MOEX bond history, T-Invest coupons и применяет ручную валидацию мэппинга, если файл `additional_public_data_manual_review_validated.xlsx` уже лежит в `data_processed/`.

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_public_market.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/public_market_data_full.xlsx"
```

Если нужно быстро пересобрать только итоговый workbook из уже скачанных промежуточных файлов:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_public_market.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/public_market_data_full.xlsx" \
  --skip-download
```

Если MOEX/CBR/World Bank уже скачаны, а нужно переобновить только T-Invest из-за токена или SSL:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 main_public_market.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/public_market_data_full.xlsx" \
  --skip-download \
  --refresh-tinvest \
  --disable-tinvest-ssl-verify
```

## Запуск 2. СПАРК DOCX

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
python main_spark.py \
  --input-dir "/Users/grigorijkrasovickij/4 крус/Диплом" \
  --pattern "СПАРК-Отчет_*.docx" \
  --quarter-grid-start 2014Q3 \
  --quarter-grid-end 2025Q4 \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_combined_report.xlsx"
```

## Запуск 3. Применить ручную валидацию мэппинга

После проверки спорных совпадений в `additional_public_data_manual_review_validated.xlsx` можно собрать отдельный workbook, где решения ручной проверки применены к MOEX / T-Invest слоям:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
PYTHONPATH=vendor python3 src/apply_manual_validation.py \
  --main-workbook "/Users/grigorijkrasovickij/Documents/Playground/data_processed/additional_public_data.xlsx" \
  --validation-workbook "/Users/grigorijkrasovickij/Documents/Playground/data_processed/additional_public_data_manual_review_validated.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/additional_public_data_validated.xlsx"
```

Для анализа после ручной проверки используйте прежде всего лист `issuer_mapping_validated` в `additional_public_data_validated.xlsx`.

## Что лежит в `additional_public_data.xlsx`

- `companies_master` — нормализованный shortlist компаний.
- `moex_instruments` — все кандидаты инструментов из MOEX ISS.
- `moex_bonds` — выбранные облигации MOEX.
- `tinvest_instruments` — все кандидаты инструментов из T-Invest.
- `tinvest_bonds` — выбранные облигации T-Invest.
- `tinvest_bond_coupons` — купонный календарь из T-Invest, если включен флаг `--tinvest-coupons`.
- `macro_cbr` — ключевая ставка и инфляция из Банка России.
- `commodity_prices_monthly` — месячные commodity prices из World Bank Pink Sheet.
- `commodity_prices_annual` — годовые commodity prices из World Bank Pink Sheet.
- `mapping_log` — лог сопоставления компаний с инструментами.
- `download_log` — технический лог загрузок и пропусков.

## Что лежит в `additional_public_data_validated.xlsx`

- `validation_summary` — компактная статистика ручной проверки.
- `validated_company_actions` — все 21 проверенное решение с флагами действий.
- `issuer_mapping_validated` — очищенная таблица подтвержденных / рекомендованных инструментов после ручной проверки.
- `tinvest_instruments_validated` — T-Invest кандидаты с флагами ручной проверки.
- `tinvest_bonds_validated` — T-Invest облигации с флагами ручной проверки.
- `moex_instruments_validated` — MOEX кандидаты с флагами ручной проверки.
- `moex_bonds_validated` — MOEX облигации с флагами ручной проверки.
- `manual_validation_checks` — исходный лист `validated_checks` из файла ручной проверки.

## Что лежит в `public_market_data_full.xlsx`

- `public_market_summary` — список листов и размеры таблиц.
- `companies_master` — нормализованный shortlist компаний.
- `firm_market_flags` — компактные флаги доступа к публичному рынку: есть ли подтвержденные акции/облигации, сколько MOEX/T-Invest совпадений найдено.
- `issuer_mapping_validated` — очищенный слой `company -> instrument` после ручной проверки.
- `moex_instruments_validated`, `moex_bonds_validated`, `moex_bond_history`, `moex_share_history` — MOEX инструменты, облигации и история торгов/доходностей, если включена.
- `tinvest_instruments_validated`, `tinvest_bonds_validated`, `tinvest_bond_coupons` — T-Invest инструменты, облигации и купонный календарь, если включен.
- `macro_cbr` — ключевая ставка и инфляция.
- `commodity_prices_monthly`, `commodity_prices_annual` — World Bank Pink Sheet: нефть, газ, уголь, железная руда и металлы.
- `mapping_log`, `download_log` — контроль качества загрузки и сопоставлений.

## Что лежит в `spark_combined_report.xlsx`

- `input_inventory` — список всех найденных файлов СПАРК в папке, включая неподдержанные форматы.
- `raw_long` — все извлеченные строки в длинном формате.
- `panel_core` — панель по ключевым финансовым метрикам.
- `panel_quarterly` — квартальная сетка по всем компаниям и кварталам заданного диапазона; пропуски остаются пустыми.
- `coverage` — покрытие форм по компании и периоду.
- `parse_log` — лог парсинга DOCX.

Если в папке со СПАРК лежат `pdf`, текущий парсер их не читает, но добавляет в `input_inventory` и пишет в `parse_log`, что файл пропущен как неподдержанный.

## Официальные источники

- [MOEX ISS reference](https://iss.moex.com/iss/reference/)
- [T-Bank token docs](https://developer.tbank.ru/invest/intro/intro/token)
- [T-Bank REST protocol](https://developer.tbank.ru/invest/intro/developer/protocols/restapi)
- [T-Invest Instruments methods](https://tinkoff.github.io/investAPI/instruments/)
- [Bank of Russia Key Rate](https://www.cbr.ru/hd_base/KeyRate/)
- [Bank of Russia Inflation](https://www.cbr.ru/hd_base/infl/)
- [World Bank Commodity Markets](https://www.worldbank.org/en/research/commodity-markets)
