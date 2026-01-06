# Export Spec – Pokédex Data Pipeline

## 0. Purpose
This document defines how derived domain models are exported into a deterministic Excel workbook `pokedata.xlsx`.

The export layer:
- takes **only derived models** (outputs of the transform layer),
- writes a workbook with defined sheets/columns,
- performs **no fetching** and **no transformation business logic**,
- may apply minimal usability formatting (freeze header row, autofilter, basic column sizing).

---

## 1. Inputs

### 1.1 Derived Inputs (from Transform Layer)
Export consumes the following derived models (see `transform.spec.md` and `data-contract.md`):
- PokemonForm
- LearnsetEntry
- Move
- Item
- Ability
- Nature
- EvolutionEdge
- TypeChart
- Meta

If some entities are missing (e.g., not fetched yet), export may either:
- fail fast with a clear error, OR
- generate partial workbook (MVP mode)
This behavior must be explicit in implementation. Default: **fail fast in production**.

---

## 2. Output

### 2.1 Output file
- Path: `data/export/pokedata.xlsx`

### 2.2 Workbook requirements
- Deterministic sheet order
- Deterministic column order within each sheet
- Deterministic row ordering (as defined below)
- All values are written as plain values (no formulas required)

### 2.3 Allowed formatting (optional)
Allowed (nice-to-have):
- Freeze top row (header)
- Auto-filter on header row
- Basic column width adjustment

Not allowed:
- Conditional formatting
- Pivot tables
- Heavy styling or UI macros

---

## 3. Sheets (Final Workbook)

Workbook contains the following sheets (in this exact order):
1. `Pokemon`
2. `Learnsets`
3. `Moves`
4. `Items`
5. `Abilities`
6. `Natures`
7. `Evolutions`
8. `TypeChart`
9. `Meta`

---

## 4. Sheet Specs

## 4.1 Sheet: Pokemon
One row = one Pokémon form.

### Columns (in order)
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
- `MOVESET_<VG1>` (string|null; semicolon-separated move keys)
- `MOVESET_<VG2>` ...
- (one `MOVESET_` column per configured version group)

### Moveset formatting
- Moveset cell is a single string with move keys separated by `;`
- No trailing delimiter
- Sorted alphabetically by move key (deterministic)

### Row ordering
- Order by `DEX_ID` ascending, then `FORM_KEY` ascending.

---

## 4.2 Sheet: Learnsets
One row = one LearnsetEntry (detailed learn method record).

### Columns (in order)
- `FORM_KEY` (string)
- `DISPLAY_NAME` (string)
- `VERSION_GROUP` (string)
- `MOVE_KEY` (string)
- `METHOD` (string)
- `LEVEL` (int|null)

### Row ordering
- Order by `FORM_KEY`, `VERSION_GROUP`, `MOVE_KEY`, `METHOD`, `LEVEL`

---

## 4.3 Sheet: Moves
One row = one move.

### Columns (in order)
- `MOVE_KEY` (string)
- `DISPLAY_NAME` (string)
- `TYPE` (string)
- `CATEGORY` (string)
- `POWER` (int|null)
- `ACCURACY` (int|null)
- `PP` (int|null)
- `PRIORITY` (int)
- `EFFECT_SHORT` (string|null)

### Row ordering
- Order by `MOVE_KEY` ascending.

---

## 4.4 Sheet: Items
One row = one item.

### Columns (in order)
- `ITEM_KEY` (string)
- `DISPLAY_NAME` (string)
- `CATEGORY` (string|null)
- `EFFECT_SHORT` (string|null)

### Row ordering
- Order by `ITEM_KEY` ascending.

---

## 4.5 Sheet: Abilities
One row = one ability.

### Columns (in order)
- `ABILITY_KEY` (string)
- `DISPLAY_NAME` (string)
- `EFFECT_SHORT` (string|null)

### Row ordering
- Order by `ABILITY_KEY` ascending.

---

## 4.6 Sheet: Natures
One row = one nature.

### Columns (in order)
- `NATURE_KEY` (string)
- `DISPLAY_NAME` (string)
- `INCREASED_STAT` (string|null)
- `DECREASED_STAT` (string|null)

### Row ordering
- Order by `NATURE_KEY` ascending.

---

## 4.7 Sheet: Evolutions
One row = one EvolutionEdge (flattened evolution condition).

### Columns (in order)
- `FROM_SPECIES_KEY` (string)
- `TO_SPECIES_KEY` (string)
- `TRIGGER` (string|null)
- `MIN_LEVEL` (int|null)
- `ITEM` (string|null)
- `TIME_OF_DAY` (string|null)
- `KNOWN_MOVE` (string|null)
- `KNOWN_MOVE_TYPE` (string|null)
- `GENDER` (int|null)
- `LOCATION` (string|null)
- `HELD_ITEM` (string|null)

### Row ordering
- Order by `FROM_SPECIES_KEY`, `TO_SPECIES_KEY`, `TRIGGER`

---

## 4.8 Sheet: TypeChart

### Representation B: Matrix
Columns (in order):
- `ATTACKING_TYPE` (string - `DISPLAY_NAME`)
- `DEFENDING_TYPE` (string - `DISPLAY_NAME`)
- `MULTIPLIER` (float)

Row ordering:
- Order `ATTACKING_TYPE`, `DEFENDING_TYPE` by config/typesOrder.json typesOrder
- IgnoreRest: by config/typesOrder.json ignoreRest (if false, then use IgnoreRestFalseMode)

(If using matrix representation, it must be defined explicitly in implementation docs.)

---

## 4.9 Sheet: Meta
Traceability and reproducibility information.

### Columns (in order)
- `KEY`
- `VALUE`

### Required keys
- `generated_at`
- `source`
- `pokeapi_base_url`
- `about_language`
- `version_groups`
- `pipeline_version`

---

### 4.10 Sheet: Assets
- create sheet `Assets`
- load data from `config/assets.csv` (`name;address`)
- rows are grouped by headings (e.g. `IMAGES:`)
- each heading starts a new column block
- all column blocks start at row 1
- at the end add new column block
- Row 1: GAMES
- Source of truth (IDs): load all game ids from config/config.json → version-groups (preserve order as defined there)
- Label mapping: for each game id, resolve label via config/gamesMap.csv (slug → label)
- Fallback rules:
    if config/gamesMap.csv is missing / unreadable → use the label from config/config.json (version-groups[].name or equivalent display field)
    if gamesMap.csv exists but no matching slug row (or label empty) → use the label from config/config.json for that specific game id
- Output: render a single-column list under GAMES using resolved labels (one per row)
- Add 'All' at the end of the list

---

## 5. Export Modes

### 5.1 MVP mode
MVP mode may export only:
- `Pokemon`
- `Meta`

### 5.2 Production mode
Production mode must export all sheets listed in section 3 and must fail fast if required inputs are missing.

---

## 6. Non-Goals
- No fetching
- No transformation business rules
- No Excel UI or macros
- No damage calc / teambuilder logic

Those belong to other layers.
