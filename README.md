# Diploma Research Pipeline

Репозиторий для двух практических задач диплома:

1. сборка общедоступных финансовых и рыночных данных по shortlist компаний;
2. парсинг `.docx`-выгрузок СПАРК в единую панель.

## Структура репозитория

- [main_public_data.py](/Users/grigorijkrasovickij/Documents/Playground/main_public_data.py) — основной entrypoint для публичных источников.
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

## Запуск 2. СПАРК DOCX

```bash
cd /Users/grigorijkrasovickij/Documents/Playground
python main_spark.py \
  --input-dir "/Users/grigorijkrasovickij/4 крус/Диплом" \
  --pattern "СПАРК-Отчет_*.docx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/data_processed/spark_combined_report.xlsx"
```

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

## Что лежит в `spark_combined_report.xlsx`

- `raw_long` — все извлеченные строки в длинном формате.
- `panel_core` — панель по ключевым финансовым метрикам.
- `coverage` — покрытие форм по компании и периоду.
- `parse_log` — лог парсинга DOCX.

## Официальные источники

- [MOEX ISS reference](https://iss.moex.com/iss/reference/)
- [T-Bank token docs](https://developer.tbank.ru/invest/intro/intro/token)
- [T-Bank REST protocol](https://developer.tbank.ru/invest/intro/developer/protocols/restapi)
- [T-Invest Instruments methods](https://tinkoff.github.io/investAPI/instruments/)
- [Bank of Russia Key Rate](https://www.cbr.ru/hd_base/KeyRate/)
- [Bank of Russia Inflation](https://www.cbr.ru/hd_base/infl/)
- [World Bank Commodity Markets](https://www.worldbank.org/en/research/commodity-markets)
