# Fetch Prompt – Pokédex Data Pipeline

## Index
1. MVP Fetch (Initial Test)
2. Full Fetch (Sample / Test Mode)
3. Production Fetch (Complete Dataset)

---

## 1. MVP Fetch – Initial Test

### Purpose
Establish the first working milestone of the Pokédex Data Pipeline.
This step validates:
- API connectivity
- entity discovery
- file-based caching with TTL
- stable identifiers

No transformations or exports are allowed.

### Scope
Fetch and cache only:
- **RawPokemon** (`/pokemon/{form_key}`)
- **RawSpecies** (`/pokemon-species/{species_id}`)

### Constraints
- Follow strictly:
  - `spec/north-star.md`
  - `spec/data-contract.md`
  - `spec/fetch.spec.md`
- Python + **pokebase**
- Raw JSON files with `_meta` + `data`
- Atomic writes
- TTL-based cache skipping

### CLI
```bash
python -m src.main fetch pokemon --limit 20
```

### Acceptance Criteria
- Creates `data/raw/pokemon/*.json`
- Creates corresponding `data/raw/species/*.json`
- Re-run skips cached files
- `--force` refreshes data

---

## 2. Full Fetch – Sample / Test Mode

### Purpose
Fetch **all raw entities** defined in the data contract, but only in a limited sample size.
This step validates:
- full entity coverage
- reference data discovery
- performance and stability of the fetch layer

### Scope
Fetch and cache:
- RawPokemon
- RawSpecies
- RawType
- RawMove
- RawItem
- RawAbility
- RawNature
- RawEvolutionChain (optional, config-gated)

### Limit Semantics
When `--limit N` is provided:
- Discovery is limited to the **first N keys per entity**
- Limit applies independently per entity

Example:
- `--limit 20` → 20 Pokémon, 20 moves, 20 items, etc.

### CLI
```bash
python -m src.main fetch all --limit 20
```

### Constraints
- Same cache, TTL, atomic write rules as MVP
- 404s must be logged and skipped
- No transforms or exports

### Acceptance Criteria
- Raw files exist for all entity folders
- Counts roughly match limit semantics
- Re-run results mostly in cache hits

---

## 3. Production Fetch – Complete Dataset

### Purpose
Perform a **complete, authoritative fetch** of all PokéAPI data required by the pipeline.
This dataset becomes the stable base for transform and export steps.

### Scope
Fetch and cache:
- All Pokémon forms
- All species
- All reference entities (types, moves, items, abilities, natures)
- All evolution chains (if enabled)

### Constraints
- No discovery limits
- Full TTL enforcement
- Safe retries and backoff
- Deterministic, restartable runs

### CLI
```bash
python -m src.main fetch all
```

Optional:
```bash
python -m src.main fetch all --force
```

### Acceptance Criteria
- Complete raw cache under `data/raw/`
- No missing mandatory entities
- Re-runs only refresh stale entries
- Dataset ready for transform/export stages

---

## Non-Goals (All Stages)
- No unit conversion
- No naming rules
- No Excel export
- No derived models

All such logic belongs to later transform/export prompts.
