# Pokédex Data Pipeline – Fetch Spec

## 0. Purpose
This spec defines **how raw data is fetched and cached** from PokéAPI to local JSON files according to `data-contract.md`.

This layer:
- Fetches and stores **raw API payloads** (plus `_meta`).
- Applies **cache rules** (skip / refresh).
- Does **not** perform domain transformations (naming, units, export formatting).

---

## 1. Data Sources & Client

### 1.1 Primary source
- **PokéAPI v2** (base URL configurable; default `https://pokeapi.co/api/v2`)

### 1.2 Fetcher
- Python fetcher uses **pokebase** as the primary API wrapper.

### 1.3 Optional HTTP-level cache (non-authoritative)
- `requests-cache` MAY be enabled to reduce repeated HTTP calls during development.
- HTTP cache is **not** a replacement for file-based raw cache.
- File-based raw cache is the **source of truth** for offline rebuilds.

---

## 2. Configuration (Input Parameters)

All fetch behavior is driven by `config/config.json` (or equivalent).

### 2.1 Required config fields
- `pokeapi_base_url`: string
- `ttl_days`:
  - `pokemon`: int
  - `species`: int
  - `evolution_chain`: int
  - `move`: int
  - `item`: int
  - `ability`: int
  - `nature`: int
  - `type`: int
- `about_language`: string (e.g. `"en"`) (used later in transform, stored in Meta sheet)
- `version_groups`: string[] (e.g. `["scarlet-violet", "sword-shield"]`) (used later in transform)
- `include_evolutions`: boolean (default true)

### 2.2 Optional config fields
- `max_retries`: int (default 5)
- `retry_backoff_seconds`: float (default 1.0; exponential backoff)
- `request_delay_seconds`: float (default 0.1) (politeness delay)
- `max_workers`: int (default 1; concurrency optional)
- `force_refresh`: boolean (default false; overrides TTL)

---

## 3. Local Raw Cache Rules

### 3.1 Storage format (per `data-contract.md`)
Each raw file is a JSON object:
- `_meta`: object
  - `fetched_at`: ISO datetime (UTC)
  - `url`: string
  - `etag`: string|null (if available / provided by library)
  - `status`: int (HTTP status if known)
- `data`: object (raw API payload)

### 3.2 Cache directory structure
Raw files are stored under `data/raw/<entity>/`:

- `data/raw/pokemon/{form_key}.json`
- `data/raw/species/{species_id}.json`
- `data/raw/evolution-chain/{evolution_chain_id}.json`
- `data/raw/move/{move_key}.json`
- `data/raw/item/{item_key}.json`
- `data/raw/ability/{ability_key}.json`
- `data/raw/nature/{nature_key}.json`
- `data/raw/type/{type_key}.json`

### 3.3 Cache decision algorithm
For a given entity record:
- If `force_refresh=true`: fetch and overwrite.
- Else if file does not exist: fetch and write.
- Else if file exists and is **stale**: fetch and overwrite.
- Else: skip fetch.

A file is **stale** if:
- `now - _meta.fetched_at > ttl_days[entity]`.

### 3.4 Atomic writes
Raw file writes MUST be atomic:
- write to temp file in same directory
- replace target file (rename)

### 3.5 Safe filenames
`form_key`, `move_key`, etc. MUST be saved as filesystem-safe filenames:
- allowed: `a-z`, `0-9`, `-`, `_`
- all other characters replaced with `_`

---

## 4. What Gets Fetched (Coverage)

### 4.1 Pokémon forms (RawPokemon)
Goal: fetch all available Pokémon forms accessible via `/pokemon` resources.

**Discovery:**
- Fetch the full list via: `GET /pokemon?limit=100000&offset=0`
- The list provides `name` keys. Each `name` is a `form_key`.

**Per record:**
- Fetch `GET /pokemon/{form_key}`
- Store as `data/raw/pokemon/{form_key}.json`

### 4.2 Species (RawSpecies)
Species data is required for:
- dex id mapping
- localized names
- flavor text (`ABOUT`)
- evolution chain reference

**Per record:**
- For each RawPokemon fetched/loaded, extract `species.url`.
- Resolve `species_id` from the URL.
- Fetch `GET /pokemon-species/{species_id}`
- Store as `data/raw/species/{species_id}.json`

### 4.3 Evolution chains (RawEvolutionChain) (optional)
If `include_evolutions=true`:

**Discovery:**
- From each RawSpecies, extract `evolution_chain.url` when present.
- Resolve `evolution_chain_id` from the URL.

**Per record:**
- Fetch `GET /evolution-chain/{evolution_chain_id}`
- Store as `data/raw/evolution-chain/{evolution_chain_id}.json`

### 4.4 Types (RawType)
Type data is required to build the TypeChart.

**Discovery:**
- Fetch list via `GET /type?limit=1000`
- Extract each `type_key`

**Per record:**
- Fetch `GET /type/{type_key}`
- Store as `data/raw/type/{type_key}.json`

### 4.5 Moves, Items, Abilities, Natures (RawMove/RawItem/RawAbility/RawNature)
These are needed to populate reference sheets.

**Discovery:**
- Moves: `GET /move?limit=100000`
- Items: `GET /item?limit=100000`
- Abilities: `GET /ability?limit=100000`
- Natures: `GET /nature?limit=1000`

**Per record:**
- Fetch corresponding `/{resource}/{key}`
- Store under `data/raw/<resource>/{key}.json`

**Note (MVP path):**
- The MVP fetch implementation MAY start with only:
  - RawPokemon + RawSpecies (+ RawType)
- Then expand to the other resources.

---

## 5. Performance & Politeness

### 5.1 Rate limiting / politeness
- Add `request_delay_seconds` delay between requests (default 0.1s).
- If PokéAPI returns rate limiting errors (e.g. 429), the fetcher MUST:
  - back off (exponential) and retry up to `max_retries`.

### 5.2 Concurrency
- Default behavior: single-threaded (`max_workers=1`) for safety.
- Concurrency MAY be enabled, but must still respect politeness delays
  (e.g. per-worker delay or global limiter).
- Regardless of concurrency, output MUST remain deterministic
  (same files and content for same inputs and same upstream).

### 5.3 Incremental runs
- Fetch runs are incremental by cache rules.
- A fetch run MUST be restartable without manual cleanup.

---

## 6. Error Handling & Observability

### 6.1 Retries
On transient failures (network errors, timeouts, 5xx, 429):
- retry up to `max_retries`
- exponential backoff using `retry_backoff_seconds`

On permanent failures (404 for a discovered key):
- log and continue (do not crash the entire run)
- mark in logs as missing

### 6.2 Logging
A fetch run MUST log:
- started_at, finished_at
- counts:
  - discovered keys per entity
  - fetched new
  - refreshed stale
  - skipped fresh
  - failed
- top error reasons (status codes)

Logs should be written to:
- console (human-readable)
- optionally `data/logs/fetch.log`

### 6.3 Run metadata
A run MAY write a summary JSON:
- `data/logs/fetch-summary.json`
containing counts and timings.

---

## 7. CLI Contract (Fetch Commands)

The pipeline exposes these commands:

### 7.1 `fetch pokemon`
- Fetch RawPokemon for all discovered `form_key`.
- Also ensures RawSpecies for all fetched/loaded pokemon.
- If `include_evolutions=true`, ensures evolution chains too.

### 7.2 `fetch reference`
- Fetch RawMove, RawItem, RawAbility, RawNature, RawType.

### 7.3 `fetch all`
- Equivalent to `fetch pokemon` + `fetch reference`.

### 7.4 Flags
- `--force`: sets `force_refresh=true`
- `--entity <name>`: restricts to a specific entity type (optional)
- `--limit <n>`: development-only limit for quick tests (optional)

---

## 8. Non-Goals (Fetch Layer)
- No transformation of units (dm/hg → m/kg).
- No display naming rules (Mega Charizard X, etc.).
- No “learnset detail shaping” beyond storing raw payloads.
- No Excel export.

All of the above belongs to transform/export specs.
