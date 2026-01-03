# Export Prompt – Pokédex Data Pipeline

## Index
1. Export MVP (Core Excel Dataset)
2. Export Extended (Gameplay Detail)
3. Export Production (Final Excel Dataset)

---

## 1. Export MVP – Core Excel Dataset

### Purpose
Create the **first usable Excel file** generated from derived data.
This step proves that transformed domain models can be exported into a structured workbook.

### Scope
Export only:
- **Pokemon** sheet
- **Meta** sheet

### Included Data
From derived models:
- PokemonForm
- Meta

### Sheet: Pokemon
One row = one Pokémon form.

Required columns:
- DEX_ID
- FORM_KEY
- DISPLAY_NAME
- FORM_GROUP
- TYPE1
- TYPE2
- HP
- ATK
- DEF
- SPA
- SPD
- SPE
- TOTAL
- ABILITY1
- ABILITY2
- HIDDEN_ABILITY
- HEIGHT_M
- WEIGHT_KG
- ABOUT
- SPRITE
- SHINY_SPRITE

No movesets, no learnsets.

### Sheet: Meta
- generated_at
- pipeline_version
- source (PokéAPI)
- about_language

### Constraints
- No Excel styling
- No formulas
- No UI logic
- Deterministic column order

### Acceptance Criteria
- `pokedata.xlsx` is generated successfully
- Workbook opens without warnings
- Data matches derived input exactly

---

## 2. Export Extended – Gameplay Detail

### Purpose
Extend the Excel export with **gameplay-relevant reference data**.

### Scope
Add sheets:
- Learnsets
- Moves
- Items
- Abilities
- Natures
- Evolutions
- TypeChart

Add moveset summary columns to `Pokemon`.

### Included Data
From derived models:
- LearnsetEntry
- Move
- Item
- Ability
- Nature
- EvolutionEdge
- TypeChart

### Pokemon Sheet Extension
Add:
- MOVESET_<VG1>
- MOVESET_<VG2>
- ...

Movesets must:
- be semicolon-separated
- be alphabetically sorted
- contain move keys only

### Constraints
- One sheet per entity
- No cross-sheet formulas
- No Excel calculations

### Acceptance Criteria
- All sheets present
- Row counts match derived data
- Referential integrity preserved by keys

---

## 3. Export Production – Final Excel Dataset

### Purpose
Generate the **final authoritative Excel dataset** for long-term use.

### Scope
Export:
- All sheets defined in data-contract.md
- Complete metadata

### Enhancements (Allowed)
- Freeze header rows
- Auto-filter headers
- Auto-size columns (basic)

### Constraints
- No business logic
- No derived calculations beyond transform layer
- Export is idempotent

### Acceptance Criteria
- Single `pokedata.xlsx` file
- Re-running export produces identical structure
- File ready for Excel-based Pokédex and teambuilder tools

---

## Non-Goals (All Stages)
- No fetching
- No transformation logic
- No Excel UI or macros

All such logic belongs to transform or frontend layers.
