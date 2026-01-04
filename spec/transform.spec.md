# Transform Spec – Pokédex Data Pipeline

## 0. Purpose
This document defines **how raw cached PokéAPI data is transformed** into normalized, derived domain models.
It sits between the fetch layer and the export layer.

This layer:
- reads **only local raw JSON files**,
- performs deterministic transformations,
- produces in-memory derived structures (or intermediate JSON for debugging),
- performs **no I/O to external services**.

No Excel formatting or UI logic belongs here.

---

## 1. Inputs

### 1.1 Raw Inputs (from Fetch Layer)
The transform layer consumes raw data exactly as defined in `data-contract.md`:

- RawPokemon
- RawSpecies
- RawEvolutionChain
- RawMove
- RawItem
- RawAbility
- RawNature
- RawType

All raw inputs:
- are loaded from `data/raw/**`
- contain `_meta` and `data`
- must be treated as immutable

Missing or partial raw data must not crash the transform; missing optional fields result in `null` outputs.

---

## 2. Outputs

### 2.1 Derived Domain Models
The transform layer produces the following derived models (as defined in `data-contract.md`):

- PokemonForm
- LearnsetEntry
- Move
- Item
- Ability
- Nature
- EvolutionEdge
- TypeChart
- Meta

Derived models:
- must be deterministic
- must be reproducible from the same raw inputs
- must not depend on execution order

---

## 3. Global Transform Rules

### 3.1 Determinism
- Same raw inputs MUST always produce identical derived outputs.
- Ordering must be stable and explicit (alphabetical by key or numeric by id).

### 3.2 Error Handling
- Missing optional fields → `null`
- Missing required raw entities:
  - log error
  - skip affected derived records
  - do not crash entire run

### 3.3 No Network Access
- Transform code MUST NOT perform any HTTP requests.

### 3.4 Version Groups
- Only version groups listed in config are processed for learnsets.
- Other version groups are ignored.

---

## 4. Transform Rules per Derived Model

## 4.1 PokemonForm

### Source
- RawPokemon
- RawSpecies

### Key Mapping
- `form_key` ← RawPokemon.name
- `dex_id` ← RawSpecies.id

### Field Rules
- `display_name`
  - Derived from RawPokemon.name
  - Apply naming rules (defined elsewhere; transform uses provided rules only)
- `form_group`
  - Classified from form_key and species flags:
    - standard
    - mega
    - gigantamax
    - regional
    - other
- `type1`, `type2`
  - Ordered by slot (slot 1 = type1)
- Base stats
  - Extract from RawPokemon.stats
  - TOTAL = sum of base stats
- Abilities
  - ability1 = first non-hidden
  - ability2 = second non-hidden (if exists)
  - hidden_ability = hidden ability (if exists)
- Height / Weight
  - height_m = RawPokemon.height / 10
  - weight_kg = RawPokemon.weight / 10
- ABOUT
  - Selected from RawSpecies.flavor_text_entries
  - Filter by language from config
  - Prefer entries matching first version-group from config
  - Fallback - use other available flavor_text_entries (in language from config)
- Sprites
  - Prefer official artwork URLs
  - Fallback to default sprites
  - If all preferred sprite sources are null, fall back to any available front-facing sprite (including animated or generation-specific sprites), prioritizing visibility over visual fidelity.
  - Same logic for shiny sprites

---

## 4.2 LearnsetEntry

### Source
- RawPokemon.moves[]

### Rules
- One LearnsetEntry per:
  - form_key
  - version_group
  - move_key
  - method
  - level
- Preserve multiple learning methods for the same move.
- Only include version groups listed in config.
- Normalize method names to lowercase strings.

---

## 4.3 Move

### Source
- RawMove

### Rules
- Copy numeric battle fields directly.
- `effect_short`:
  - extracted from effect_entries
  - filtered by language from config
  - short effect only (no flavor text)

---

## 4.4 Item

### Source
- RawItem

### Rules
- Extract category
- Extract short effect text by language
- Extract icon_url from `data.sprites.default`

---

## 4.5 Ability

### Source
- RawAbility

### Rules
- Extract short effect text by language

---

## 4.6 Nature

### Source
- RawNature

### Rules
- Extract increased_stat and decreased_stat
- Null allowed

---

## 4.7 EvolutionEdge

### Source
- RawEvolutionChain

### Rules
- Flatten evolution tree into directed edges
- One row per evolution condition
- Extract:
  - trigger
  - min_level
  - item
  - time_of_day
  - known_move
  - known_move_type
  - gender
  - location
  - held_item
- Missing conditions allowed as null

---

## 4.8 TypeChart

### Source
- RawType

### Rules
- Build complete type effectiveness matrix or relation list
- Multipliers:
  - 0, 0.5, 1, 2
- Type names must be normalized and stable

---

## 4.9 Meta

### Rules
- generated_at = current UTC time
- pipeline_version = config value
- about_language = config value
- version_groups = config value
- source = PokéAPI

---

## 5. Ordering & Stability

### 5.1 Ordering Rules
- PokemonForm: ordered by dex_id, then form_key
- LearnsetEntry: ordered by form_key, version_group, move_key, method, level
- Reference tables: ordered alphabetically by key

### 5.2 Idempotency
- Running transform multiple times without changing raw data MUST produce identical outputs.

---

## 6. Non-Goals
- No fetching
- No Excel formatting
- No UI logic
- No damage calculation
- No team validation

These belong to export or frontend layers.
