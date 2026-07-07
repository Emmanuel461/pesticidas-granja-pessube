# Análise do questionário sobre pesticidas em Granja de Pessubé

Repositório para preparar, calcular e visualizar os resultados do questionário aplicado às horticultoras de Granja de Pessubé sobre práticas de uso de pesticidas, conhecimento de segurança e perceções sobre impactos na saúde humana e no ambiente.

## Regra de arquitetura

- Python prepara os dados e calcula todos os resultados.
- A aplicação web em `docs/` apenas lê JSON pré-calculados e renderiza tabelas/gráficos.
- Os dados ficam somente em `data/`.
- Não existe `docs/data/`.

## Estrutura

```text
data/
  raw/              Excel original local.
  processed/        Base limpa, catálogos e CSV de auditoria.
  app/              JSON pré-calculados consumidos pela app.

docs/               Aplicação estática.
scripts/            Scripts Python do fluxo de trabalho.
notebooks/          Espaço para exploração opcional.
images/             Espaço para figuras exportadas, se forem necessárias depois.
```

## Fluxo de trabalho

```bash
pip install -r requirements.txt
python scripts/01_prepare_dataset.py --input data/raw/questionario_original.xlsx
python scripts/02_generate_app_outputs.py
python -m http.server 8000
```

Abrir no navegador:

```text
http://localhost:8000/docs/
```

## Scripts

### `scripts/01_prepare_dataset.py`

Limpa e organiza o Excel original. Gera:

```text
data/processed/questionario_clean.xlsx
data/processed/questionario_clean_wide.csv
data/processed/single_choice_wide.csv
data/processed/multiple_choice_long.csv
data/processed/question_catalog.csv
data/processed/question_catalog_app.csv
data/processed/value_audit.csv
data/app/questions.json
data/app/app_config.json
```

### `scripts/02_generate_app_outputs.py`

Calcula todos os resultados usados pela aplicação web. Gera:

```text
data/app/dashboard_cards.json
data/app/descriptive_results.json
data/app/multiple_choice_results.json
data/app/crosstab_results.json
data/app/app_manifest.json
```

Também gera CSV de auditoria em `data/processed/`.

## Publicação no GitHub Pages

Como a app está em `docs/` e os dados estão em `data/app/`, publicar GitHub Pages a partir da raiz do repositório e abrir:

```text
https://usuario.github.io/repositorio/docs/
```
