# Assets Spec – Pokédex Data Pipeline

## Purpose
This specification defines **icon-related derived data** that extends the existing pipeline
without modifying core fetch/transform/export specs.

It covers:
- Pokémon **Type icons** (from PokéAPI RawType sprites)
- **Item icons** (from PokéAPI RawItem sprites)

This spec is an additive extension and must be referenced explicitly by transform/export prompts.

---

## Requirements & Source of Truth

This spec MUST be implemented strictly according to:
- `spec/north-star.md`
- `spec/data-contract.md`
- `spec/assets.spec.md` (this document)

Rules:
- No fetching beyond what is already defined in fetch specs
- No scraping of external websites
- Only data available in cached raw payloads may be used
- Output must be deterministic

---

## 1. Type Icons

### Source
- `RawType.data.sprites`

### Output
- Derived model: `Type`
- Field: `icon_url`

### Selection Rule (Deterministic, Future-Proof)

1. Traverse all keys under `RawType.data.sprites` matching `generation-*`
2. Sort generation keys **lexicographically ascending**
3. Select the **last generation** (highest key, e.g. `generation-viii`)
4. Within that generation:
   - Traverse all game/version keys
   - Sort game keys lexicographically ascending
   - Select the **last available** `name_icon`
5. If no `name_icon` is found across all generations:
   - Set `icon_url = null`

### Notes
- No configuration flags are required
- New generations will be picked up automatically
- This rule must not depend on `version_groups`

---

## 2. Item Icons

### Source
- `RawItem.data.sprites.default`

### Output
- Derived model: `Item`
- Field: `icon_url`

### Rules
- If `sprites.default` exists → use as `icon_url`
- If missing → `icon_url = null`
- No version-based selection applies

---

## 3. Derived Models Impact

This spec requires the following derived models to exist in `data-contract.md`:

### Type (Derived)
- `type_key`: string
- `icon_url`: string|null

### Item (Derived)
- add field:
  - `icon_url`: string|null

No other derived models are affected.

---

## 4. Excel Export Impact

### Types Sheet
A new sheet `Types` must be exported.

Columns (in order):
- `TYPE_KEY`
- `ICON_URL`

Row ordering:
- Order by `TYPE_KEY` ascending

### Items Sheet
Add column:
- `ICON_URL`

Column position:
- Append at the end of the existing `Items` sheet

---

## Non-Goals
- No move or ability icons
- No local asset bundling
- No external CDN usage
- No UI or Excel styling logic
