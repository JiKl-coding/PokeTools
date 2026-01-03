# Development Notes (Legacy README)

This file contains the previous root `README.md` content, kept for reference.

If you’re looking for current usage docs, see:
- `manual.md`
- `spec/` (fetch/transform/export specs)

---

## Previous Root README

PokéTools Fetch MVP

Overview

- Fetches and caches raw PokéAPI data for Pokemon forms and their species.
- Uses file-based cache JSON structure with `_meta` and `data` keys.

Setup

- Ensure Python 3.10+.
- Create/activate a virtual environment and install dependencies:

```
python -m venv .venv
.venv\\Scripts\\python.exe -m ensurepip --upgrade
.venv\\Scripts\\python.exe -m pip install pokebase requests
```

Config

- Edit config/config.json to adjust `pokeapi_base_url` and TTLs.

Run

- Fetch first 20 Pokemon forms and their species (Windows):

```
.venv\\Scripts\\python.exe -m src.main fetch pokemon --limit 20
```

- Force refresh regardless of TTL:

```
.venv\\Scripts\\python.exe -m src.main fetch pokemon --limit 20 --force
```

Cache Files

- Pokemon: data/raw/pokemon/{form_key}.json
- Species: data/raw/species/{species_id}.json

Notes

- Atomic writes: temp file + rename to avoid partial writes.
- Safe filenames: normalized `form_key` to alphanumeric, dash, underscore.
