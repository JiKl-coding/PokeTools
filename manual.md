# PokeTools Manual (Simple)

This project is a 3-stage pipeline:

1) **Fetch** raw PokéAPI data into `data/raw/**`
2) **Transform** raw -> derived models into `data/derived/**`
3) **Export** derived models -> `data/export/pokedata.xlsx`

The supported way to run everything is via the single CLI entry point:

```bash
python -m src.main <command>
```

`src.main` delegates to the main CLI dispatcher in `src/fetch.py`, which currently hosts **fetch**, **transform**, and **export** subcommands.

## Index

- [Install dependencies](#install-dependencies)
- [0) Configuration](#0-configuration)
- [1) Fetch](#1-fetch-writes-dataraw)
- [2) Transform](#2-transform-reads-dataraw-writes-dataderived)
- [3) Export](#3-export-reads-dataderived-writes-dataexportpokedataxlsx)
- [File-by-file quick notes](#file-by-file-quick-notes-what-each-py-is-for)
- [Full fetch (full data)](#full-fetch-full-data)

---

## Install dependencies

If you see an error like:

```text
ModuleNotFoundError: No module named 'requests'
```

…it means you’re running with a Python environment that does not have this project’s dependencies installed yet.

From the repo root, install dependencies from `pyproject.toml`:

```bash
python -m pip install -e .
```

(Non-editable alternative: `python -m pip install .`)

If your environment can’t do editable installs, you can also install just the runtime deps:

```bash
python -m pip install requests pokebase openpyxl
```

Then retry your command (for example `python -m src.main transform production`).

## 0) Configuration

Edit `config/config.json` before running.

Common keys:
- `pokeapi_base_url` (usually `https://pokeapi.co/api/v2`)
- `ttl_days.*` (cache TTL per entity)
- `about_language` (e.g. `en`)
- `version_groups` (**important** for learnsets; if empty, learnsets will be empty)
- `include_evolutions` (used by fetch `all` to optionally fetch evolution chains)
- `force_refresh` (default fetch behavior)

---

## 1) Fetch (writes `data/raw/**`)

These commands download/cache raw PokéAPI payloads.

### Fetch Pokémon + species
```bash
python -m src.main fetch pokemon --limit 20
# optional
python -m src.main fetch pokemon --limit 20 --force
```

### Fetch reference entities (sample mode)
```bash
python -m src.main fetch reference --limit 20
```

### Fetch everything (sample mode)
```bash
python -m src.main fetch all --limit 20
```

Notes:
- Fetch uses TTL-based caching; reruns mostly skip.
- Raw data is stored under `data/raw/<entity>/...`.

---

## 2) Transform (reads `data/raw/**`, writes `data/derived/**`)

Transform never hits the network.

### Transform MVP (PokemonForm + Meta)
```bash
python -m src.main transform mvp
```

### Transform Extended (+ LearnsetEntry + EvolutionEdge)
```bash
python -m src.main transform extended
```

### Transform Production (full derived dataset)
```bash
python -m src.main transform production
```

Notes:
- If `config.version_groups` is `[]`, then **LearnsetEntry output will be empty** by design.
- Outputs are written to `data/derived/*.json`.

---

## 3) Export (reads `data/derived/**`, writes `data/export/pokedata.xlsx`)

Export never fetches and must not alter derived data.

### Export MVP (Pokemon + Meta sheets)
```bash
python -m src.main export mvp
```

### Export Extended (adds Learnsets/Moves/Items/Abilities/Natures/Evolutions/TypeChart)
```bash
python -m src.main export extended
```

### Export Production (full workbook, fail-fast)
```bash
python -m src.main export production
```

Output:
- `data/export/pokedata.xlsx`

---

## File-by-file quick notes (what each `.py` is for)

- `src/main.py`
  - The entry point used by all commands: `python -m src.main ...`

- `src/fetch.py`
  - The CLI dispatcher and fetch implementation.
  - Hosts the subcommands: `fetch`, `transform`, `export`.

- `src/transform.py`
  - Transform orchestration (MVP/Extended/Production runners).

- `src/transform_learnset.py`
  - Learnset parsing helpers used by transform.

- `src/transform_evolution.py`
  - Evolution-chain flattening helpers used by transform.

- `src/transform_reference.py`
  - Move/Item/Ability/Nature/TypeChart builders used by transform.

- `src/export.py`
  - Excel export implementation (MVP/Extended/Production export runners).

- `src/cache/io.py`
  - File cache utilities (atomic JSON writes, TTL staleness checks, safe filenames).

---

## Full fetch (full data)

1) Fetch all raw entities:

```bash
python -m src.main fetch all
```

2) Build the full derived dataset:

```bash
python -m src.main transform production
```

3) Write the full Excel workbook:

```bash
python -m src.main export production
```

Tip: if you want learnsets populated, set `version_groups` in `config/config.json` before running step 2.
