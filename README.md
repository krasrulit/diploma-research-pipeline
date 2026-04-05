# Research Workspace

Этот репозиторий я привел к формату, в котором удобно вести несколько параллельных исследовательских подпроектов и потом спокойно переносить их на GitHub.

## Где что лежит

- [projects/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/README.md) — карта всех активных подпроектов.
- [projects/public_data_pipeline/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/public_data_pipeline/README.md) — pipeline для общедоступных данных по shortlist компаний.
- [projects/spark_docx_parser/README.md](/Users/grigorijkrasovickij/Documents/Playground/projects/spark_docx_parser/README.md) — парсер DOCX-отчетов СПАРК.

## GitHub

Сейчас у локального репозитория нет подключенного `git remote`, поэтому код уже разложен в GitHub-friendly структуру, но пушить пока некуда.

Когда будете готовы, достаточно:

```bash
git remote add origin <ваш-github-url>
git checkout -b codex/public-data-pipeline
git add .
git commit -m "Set up research project structure"
git push -u origin codex/public-data-pipeline
```
