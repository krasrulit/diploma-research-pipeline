# Additional Public Data Pipeline

Подпроект для автоматической сборки общедоступных данных по shortlist компаний.

## Источники

- Moscow Exchange ISS API
- Банк России
- World Bank Pink Sheet
- optional Cbonds API

## Навигация по папкам

- [src/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/src/README.md)
- [data_raw/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/data_raw/README.md)
- [data_processed/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/data_processed/README.md)
- [notebooks/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/notebooks/README.md)

## Установка

Из корня проекта:

```bash
cd /Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline
pip install -r requirements.txt
```

## Запуск полного pipeline

```bash
cd /Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline
python main.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/data_processed/additional_public_data.xlsx"
```

## Быстрый smoke test

```bash
cd /Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline
python main.py \
  --shortlist "/Users/grigorijkrasovickij/4 крус/Диплом/Нефтегазовые_компании_shortlist.xlsx" \
  --output "/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/data_processed/additional_public_data_test.xlsx" \
  --max-companies 3
```

## Дополнительные флаги

- `--moex-history` — тянет историю торгов по облигациям с MOEX.
- `--with-cbonds` — включает optional Cbonds loader, если настроены credentials.

## Cbonds credentials

```bash
export CBONDS_API_LOGIN="..."
export CBONDS_API_PASSWORD="..."
export CBONDS_ENDPOINT_PATH="API_Emissions/get_flow/"
export CBONDS_API_URL="https://cbonds.com/api/services/json/"
```

Если переменных нет, pipeline пишет `SKIPPED` в `download_log` и продолжает работу.

## Что лежит в итоговом Excel

- `companies_master`
- `moex_instruments`
- `moex_bonds`
- `macro_cbr`
- `commodity_prices_monthly`
- `commodity_prices_annual`
- `mapping_log`
- `download_log`

## Официальные источники

- [MOEX ISS reference](https://iss.moex.com/iss/reference/)
- [MOEX security search](https://iss.moex.com/iss/securities.json)
- [MOEX bonds market](https://iss.moex.com/iss/engines/stock/markets/bonds/securities.json)
- [Bank of Russia Key Rate](https://www.cbr.ru/hd_base/KeyRate/)
- [Bank of Russia Inflation](https://www.cbr.ru/hd_base/infl/)
- [World Bank Commodity Markets](https://www.worldbank.org/en/research/commodity-markets)
