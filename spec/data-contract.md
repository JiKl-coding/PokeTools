# Pokédex Data Pipeline – Data Contract

## 0. Definitions
This document defines:
- The **raw entities** we store locally (direct PokéAPI payloads).
- The **derived entities** used to generate the final Excel dataset.
- The **Excel output schema** (sheets, columns, formats).

**No implementation details** belong here.

---

## 0.1 Locked Decisions
- `DISPLAY_NAME` naming_style: `slug_titlecase`
- `TypeChart` representation: `matrix`

---

## 1. Raw Storage Contract (Cached API Payloads)

### 1.1 RawPokemon (Form-level)
**Source:** PokéAPI `/pokemon/{name}`  
**Storage key:** `form_key` (PokéAPI slug, e.g. `charizard`, `charizard-mega-x`, `pikachu-gmax`)  
**File path:** `data/raw/pokemon/{form_key}.json`

**Stored as:**
- `_meta`: object (`fetched_at`, `url`, optional `etag`)
- `data`: object (raw API response)

**Required raw fields (must be present in `data`):**
- `id`: int
- `name`: string
- `species.name`: string
- `species.url`: string
- `types[]`: list
- `stats[]`: list
- `abilities[]`: list
- `moves[]`: list (including `version_group_details`)
- `sprites`: object
- `height`: int (decimetres)
- `weight`: int (hectograms)

**Notes:**
- One RawPokemon record represents exactly one **Pokémon form**.

---

### 1.2 RawSpecies (Species-level)
**Source:** PokéAPI `/pokemon-species/{id}`  
**Storage key:** `species_id` (int)  
**File path:** `data/raw/species/{species_id}.json`

**Required raw fields (must be present in `data`):**
- `id`: int
- `name`: string
- `names[]`: list (localized names)
- `flavor_text_entries[]`: list (localized flavor texts)
- `evolution_chain.url`: string|null

**Notes:**
- Species-level data is shared across multiple forms.

---

### 1.3 RawEvolutionChain
**Source:** PokéAPI `/evolution-chain/{id}`  
**Storage key:** `evolution_chain_id` (int)  
**File path:** `data/raw/evolution-chain/{evolution_chain_id}.json`

**Required raw fields (must be present in `data`):**
- `id`: int
- `chain`: object (evolution tree)
  - `species.name`, `species.url`
  - `evolves_to[]`
  - `evolution_details[]` (trigger and conditions)

**Notes:**
- This payload is used to derive **evolution edges** between species.
- Evolution is species-level, not form-level.

---

### 1.4 RawMove
**Source:** PokéAPI `/move/{name}`  
**Storage key:** `move_key` (slug)  
**File path:** `data/raw/move/{move_key}.json`

**Required raw fields (must be present in `data`):**
- `id`, `name`
- `type.name`
- `damage_class.name`
- `power`, `accuracy`, `pp`, `priority`
- `effect_entries[]` (localized effect text)

---

### 1.5 RawItem
**Source:** PokéAPI `/item/{name}`  
**Storage key:** `item_key` (slug)  
**File path:** `data/raw/item/{item_key}.json`

**Required raw fields:**
- `id`, `name`
- `category.name`
- `effect_entries[]` (localized effect text)

---

### 1.6 RawAbility
**Source:** PokéAPI `/ability/{name}`  
**Storage key:** `ability_key` (slug)  
**File path:** `data/raw/ability/{ability_key}.json`

**Required raw fields:**
- `id`, `name`
- `effect_entries[]` (localized effect text)

---

### 1.7 RawNature
**Source:** PokéAPI `/nature/{name}`  
**Storage key:** `nature_key` (slug)  
**File path:** `data/raw/nature/{nature_key}.json`

**Required raw fields:**
- `id`, `name`
- `increased_stat.name` (nullable)
- `decreased_stat.name` (nullable)

---

### 1.8 RawType
**Source:** PokéAPI `/type/{name}`  
**Storage key:** `type_key` (slug)  
**File path:** `data/raw/type/{type_key}.json`

**Required raw fields:**
- `id`, `name`
- `damage_relations`:
  - `double_damage_to[]`, `half_damage_to[]`, `no_damage_to[]`
  - `double_damage_from[]`, `half_damage_from[]`, `no_damage_from[]`

---

## 2. Derived Data Contract (Normalized Domain Models)

### 2.1 PokemonForm (Derived)
Represents one Pokémon form to be exported as one row in the Excel Pokédex sheet.

**Primary keys:**
- `form_key`: string (PokéAPI slug)
- `dex_id`: int (National Dex species id)

**Fields:**
- `dex_id`: int
- `form_key`: string
- `display_name`: string
- `form_group`: enum string (e.g. `Standard`, `Mega`, `Gigantamax`, `Regional`, `Other`)
- `type1`: string
- `type2`: string|null
- `base_hp`: int
- `base_atk`: int
- `base_def`: int
- `base_spa`: int
- `base_spd`: int
- `base_spe`: int
- `base_total`: int
- `ability1`: string|null
- `ability2`: string|null
- `hidden_ability`: string|null
- `about`: string|null
- `sprite_url`: string|null
- `shiny_sprite_url`: string|null
- `height_m`: float|null
- `weight_kg`: float|null

**Sprites (selection rule):**
- Prefer official artwork when available; otherwise fallback to default front sprite.
- Apply the same rule for shiny.

**Display name rule (`naming_style = slug_titlecase`):**
- Input: `form_key` (e.g. `charizard-mega-y`)
- Replace `-` with spaces
- Split into tokens by spaces
- Title Case each token
- Single-letter tokens are uppercase (`x`→`X`, `y`→`Y`)
- Example: `charizard-mega-y` → `Charizard Mega Y`

---

### 2.2 LearnsetEntry (Derived)
Represents one learnable move for a given Pokémon form in a specific version group.

**Primary keys (composite):**
- `form_key`: string
- `display_name`: string
- `version_group`: string
- `move_key`: string
- `method`: string
- `level`: int|null

**Fields:**
- `form_key`: string
- `display_name`: string
- `version_group`: string
- `move_key`: string
- `method`: enum string (e.g. `level-up`, `machine`, `egg`, `tutor`, `stadium-surfing-pikachu`, etc.)
- `level`: int|null (only for `level-up` when known)

**Notes:**
- One row equals one learn record. The same move may appear multiple times for a form+VG if learned by multiple methods.

---

### 2.3 Move (Derived)
**Primary key:** `move_key`  
**Fields:**
- `move_key`: string
- `display_name`: string
- `type`: string
- `category`: string (`physical` / `special` / `status`)
- `power`: int|null
- `accuracy`: int|null
- `pp`: int|null
- `priority`: int
- `effect_short`: string|null

---

### 2.4 Item (Derived)
**Primary key:** `item_key`  
**Fields:**
- `item_key`: string
- `display_name`: string
- `category`: string|null
- `effect_short`: string|null
- `icon_url` : string|null

---

### 2.5 Ability (Derived)
**Primary key:** `ability_key`  
**Fields:**
- `ability_key`: string
- `display_name`: string
- `effect_short`: string|null

---

### 2.6 Nature (Derived)
**Primary key:** `nature_key`  
**Fields:**
- `nature_key`: string
- `display_name`: string
- `increased_stat`: string|null
- `decreased_stat`: string|null

---

### 2.7 EvolutionEdge (Derived)
Represents one directed evolution relationship between two species.

**Primary keys (composite):**
- `from_dex_id`: int
- `to_dex_id`: int
- `trigger`: string
- `min_level`: int|null
- `item_key`: string|null

**Fields:**
- `from_dex_id`: int
- `to_dex_id`: int
- `trigger`: string (e.g. `level-up`, `use-item`, `trade`, ...)
- `min_level`: int|null
- `item_key`: string|null
- `time_of_day`: string|null
- `min_happiness`: int|null
- `known_move_key`: string|null
- `known_move_type`: string|null
- `location`: string|null
- `gender`: int|null
- `held_item_key`: string|null

**Notes:**
- This model is derived by flattening the evolution-chain tree.
- Not all conditions may be populated; missing values remain null.

---

### 2.8 TypeChart (Derived) – Matrix (locked)
- A 2D matrix of multipliers: `attacking_type` x `defending_type` → `multiplier` (0, 0.5, 1, 2)
- Rows/columns are ordered by `type_key` ascending (deterministic)

### 2.9 X Type (Derived)
**Primary key:** `type_key`

**Fields:**
- `type_key`: string
- `display_name`: string
- `icon_url`: string|null

---

## 3. Excel Output Contract

### 3.1 Output file
- File path: `data/export/pokedata.xlsx`

### 3.2 Sheets
The workbook contains the following sheets:
- `Pokemon`
- `Learnsets`
- `Moves`
- `Items`
- `Abilities`
- `Natures`
- `Evolutions`
- `TypeChart`
- `Meta`

---

### 3.3 Sheet: Pokemon
**One row = one Pokémon form.**

**Columns (in order):**
- `DEX_ID` (int)
- `FORM_KEY` (string)
- `DISPLAY_NAME` (string)
- `FORM_GROUP` (string)
- `TYPE1` (string)
- `TYPE2` (string|null)
- `HP` (int)
- `ATK` (int)
- `DEF` (int)
- `SPA` (int)
- `SPD` (int)
- `SPE` (int)
- `TOTAL` (int)
- `ABILITY1` (string|null)
- `ABILITY2` (string|null)
- `HIDDEN_ABILITY` (string|null)
- `HEIGHT_M` (float|null)
- `WEIGHT_KG` (float|null)
- `ABOUT` (string|null)
- `SPRITE` (string|null)
- `SHINY_SPRITE` (string|null)
- `MOVESET_<VG1>` (string|null; semicolon-separated)
- `MOVESET_<VG2>` ...
- ...

**Moveset format (summary):**
- A single string containing move keys separated by `;`
- No trailing delimiter
- Sorted order: **alphabetical** by `move_key`

---

### 3.4 Sheet: Learnsets
**One row = one learn record (detailed learnset).**

**Columns (in order):**
- `FORM_KEY` (string)
- `DISPLAY_NAME` (string)
- `VERSION_GROUP` (string)
- `MOVE_KEY` (string)
- `METHOD` (string)
- `LEVEL` (int|null)

---

### 3.5 Sheet: Moves
**Columns:**
- `MOVE_KEY`
- `DISPLAY_NAME` (string)
- `TYPE`
- `CATEGORY`
- `POWER`
- `ACCURACY`
- `PP`
- `PRIORITY`
- `EFFECT_SHORT`

---

### 3.6 Sheet: Items
**Columns:**
- `ITEM_KEY`
- `DISPLAY_NAME` (string)
- `CATEGORY`
- `EFFECT_SHORT`

---

### 3.7 Sheet: Abilities
**Columns:**
- `ABILITY_KEY`
- `DISPLAY_NAME` (string)
- `EFFECT_SHORT`

---

### 3.8 Sheet: Natures
**Columns:**
- `NATURE_KEY`
- `DISPLAY_NAME` (string)
- `INCREASED_STAT`
- `DECREASED_STAT`

---

### 3.9 Sheet: Evolutions
**One row = one evolution edge (species-level).**

**Columns (in order):**
- `FROM_DEX_ID` (int)
- `TO_DEX_ID` (int)
- `TRIGGER` (string)
- `MIN_LEVEL` (int|null)
- `ITEM_KEY` (string|null)
- `TIME_OF_DAY` (string|null)
- `MIN_HAPPINESS` (int|null)
- `KNOWN_MOVE_KEY` (string|null)
- `KNOWN_MOVE_TYPE` (string|null)
- `LOCATION` (string|null)
- `GENDER` (int|null)
- `HELD_ITEM_KEY` (string|null)

---

### 3.10 Sheet: TypeChart (Matrix)
- Columns: `ATTACKING_TYPE`, then one column per defending type (ordered by type_key asc)
- Rows: one row per attacking type (ordered by type_key asc)
- Cells: multiplier (0, 0.5, 1, 2)
- Use `DISPLAY_NAME` instead of `TYPE_KEY`

---

### 3.11 Sheet: Types
- `DISPLAY_NAME` (string)
- `TYPE_KEY`
- `ICON_URL`

---

### 3.12 Sheet: Meta
**Purpose:** traceability and reproducibility.

**Fields (suggested):**

---


---


---

## 2.X GameVersions Availability Lists (Derived, Export-only)

This project materializes **game version group availability** as dedicated **Excel lists**
in the `GAMEVERSIONS` sheet.

Availability is **not** stored as per-entity flags (e.g. `available_in_vgs`) on derived
domain models (PokemonForm, Move, Ability, Item). This avoids redundancy and keeps
domain models stable.

### Purpose
- Provide fast, precomputed lists for Excel Data Validation dropdowns and VBA.
- Avoid deriving availability dynamically in client code.

### List values
Each availability list contains **display names** (not keys/slugs).

**Assumption (locked):**
- Within each list, `display_name` values are unique (no duplicates).
- Client lookups are performed by `display_name`.

If uniqueness is ever violated in the future, the contract must be revised to include
keys (e.g. `display_name (key)`), but for the current dataset this is not necessary.

---
