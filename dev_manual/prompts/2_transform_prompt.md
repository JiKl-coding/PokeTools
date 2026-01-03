# Transform Prompt – Pokédex Data Pipeline

## Index
1. Transform MVP (Core Pokédex)
2. Transform Extended (Learnsets & Evolutions)
3. Transform Production (Full Dataset)

---

## 1. Transform MVP – Core Pokédex

### Purpose
Create the **first usable, structured dataset** derived from raw PokéAPI payloads.
This step proves that raw data can be reliably transformed into a Pokédex-ready format.

### Scope
Transform raw data into:
- **PokemonForm** (derived)
- Minimal **Meta** information

### Included Data
From raw cache:
- RawPokemon
- RawSpecies

Derived output:
- Basic form identity
- Types
- Base stats + total
- Abilities (1, 2, hidden)
- Height (m) and weight (kg)
- Sprite and shiny sprite URLs
- ABOUT text (language-selected)
- Display name
- Form group classification

### Explicitly Excluded
- Learnsets
- Moves reference data
- Items, abilities reference tables
- Evolutions
- Excel export

### Output
- In-memory derived objects
- Optional JSON debug output (development only)

### Acceptance Criteria
- Each Pokémon form produces exactly one derived record
- Base stats sum correctly to TOTAL
- Missing optional fields handled as null
- No crashes on missing raw fields

---

## 2. Transform Extended – Learnsets & Evolutions

### Purpose
Extend the Pokédex data with **gameplay-relevant detail** required for teambuilding and validation.

### Scope
Add transformation for:
- **LearnsetEntry** (detailed learnsets)
- Moveset summaries per version group
- **EvolutionEdge** (species-level)

### Included Data
From raw cache:
- RawPokemon (moves + version_group_details)
- RawMove (reference)
- RawSpecies
- RawEvolutionChain

Derived output:
- LearnsetEntry rows:
  - form_key
  - version_group
  - move_key
  - method
  - level
- Moveset summary strings per version group
- Evolution edges with triggers and conditions

### Explicitly Excluded
- Damage calculation logic
- Team legality rules
- Excel UI logic

### Acceptance Criteria
- Learnsets are complete and deduplicated
- Multiple learn methods for same move preserved
- Evolution chains flattened correctly
- Version group filtering respected

---

## 3. Transform Production – Full Dataset

### Purpose
Produce the **final authoritative derived dataset** ready for Excel export and long-term use.

### Scope
Transform and normalize:
- PokemonForm
- LearnsetEntry
- Move
- Item
- Ability
- Nature
- TypeChart
- EvolutionEdge
- Meta

### Included Data
From raw cache:
- All raw entities defined in data-contract.md

Derived output:
- Fully normalized domain models
- Stable keys and references
- Deterministic ordering for reproducibility

### Constraints
- No network access
- Pure transformation only
- Idempotent output (same input → same output)

### Acceptance Criteria
- Full dataset builds without errors
- Referential integrity preserved
- Ready for Excel export without further shaping

---

## Non-Goals (All Stages)
- No fetching
- No Excel formatting or styling
- No UI or frontend logic

These belong to export or UI layers.
