# Assets Prompt – Pokédex Data Pipeline

## Requirements & Source of Truth

This prompt MUST be implemented strictly according to the following documents:

- `spec/north-star.md`
- `spec/data-contract.md`
- `spec/assets.spec.md`

Rules:
- Derived models MUST conform exactly to `data-contract.md`
- Icon selection rules are fully defined in `assets.spec.md`
- No fetching, scraping, or network access is allowed
- Only cached raw data may be used
- Output must be deterministic

---

## Index
1. Assets MVP (Types + Item Icons)
2. Assets Production (Complete Icon Dataset)

---

## 1. Assets MVP – Types + Item Icons

### Purpose
Generate the **first usable icon dataset** for Excel and UI usage.
This step validates:
- correct extraction of type icons from RawType
- correct extraction of item icons from RawItem
- deterministic icon selection

### Scope
Generate derived data for:
- **Type**
  - `type_key`
  - `icon_url`
- **Item**
  - `item_key`
  - `icon_url`

### Included Data
From raw cache:
- RawType
- RawItem

### Explicitly Excluded
- Move icons
- Ability icons
- Any local asset bundling
- Excel export

### Output
- In-memory derived objects:
  - `Type[]`
  - updated `Item[]` with `icon_url`
- Optional JSON debug output (development only)

### Acceptance Criteria
- One Type record per Pokémon type
- `icon_url` selected according to `assets.spec.md`
- Items have `icon_url` when available, otherwise null
- No crashes on missing sprite data

---

## 2. Assets Production – Complete Icon Dataset

### Purpose
Generate the **final authoritative icon dataset** used by export and UI layers.

### Scope
Generate:
- All Type icons (latest available)
- All Item icons

### Constraints
- No discovery limits
- No network access
- Pure transformation only
- Idempotent output

### Output
- Fully populated derived models ready for export

### Acceptance Criteria
- All types resolved deterministically
- All items processed
- Output matches `assets.spec.md` exactly
- Safe to consume by export layer

---

## Non-Goals (All Stages)
- No fetching
- No transformation of unrelated entities
- No Excel writing

Those concerns belong to other pipeline stages.
