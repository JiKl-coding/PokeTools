# PokeTools

PokeTools is a small, spec-driven Pokédex data pipeline:

1) **Fetch** raw PokéAPI payloads into `data/raw/**` (cached with TTL)
2) **Transform** raw JSON into derived datasets in `data/derived/**` (offline, deterministic)
3) **Export** derived datasets to `data/export/pokedata.xlsx`

Thanks to **PokéAPI** (data source) and **pokebase** (Python client).

## Quick start

1) Install dependencies (from repo root):

```bash
python -m pip install -e .
```

2) Configure the pipeline:
- Edit `config/config.json`
- If you want learnsets populated, set `version_groups` (if empty, learnsets will be empty by design)

3) Run the pipeline:

```bash
python -m src.main fetch all
python -m src.main transform production
python -m src.main export production
```

Outputs:
- Raw cache: `data/raw/**`
- Derived JSON: `data/derived/**`
- Excel workbook: `data/export/pokedata.xlsx`

## More docs

- Full usage guide: `manual.md`
- Specs: `spec/` (fetch/transform/export + data contracts)
- Legacy notes: `development_notes/README.md`
