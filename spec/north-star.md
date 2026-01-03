# Pokédex Data Pipeline – North Star

## 1. Purpose
Describe **why this project exists**.

- Build a reliable, offline Pokémon dataset for Excel-based analysis and tooling.
- Eliminate manual data maintenance when new generations, forms, or updates are released.
- Provide a single, reproducible data pipeline from PokéAPI to Excel.

This project focuses strictly on **data generation**, not on UI or gameplay logic.

---

## 2. Scope
Define **what the project does and does not do**.

### In scope
- Fetch raw Pokémon data from PokéAPI.
- Cache raw API responses locally.
- Transform raw data into a structured dataset.
- Export the dataset into an Excel-compatible format (`.xlsx`).
- Support Pokémon forms (Mega, G-Max, regional variants, etc.).
- Support move learnsets per selected version groups.

### Out of scope
- Any graphical or interactive frontend (handled manually in Excel).
- Battle simulation, damage calculation, or rules enforcement.
- Competitive meta analysis or legality validation.
- Persistent user data storage (teams, notes, history).

---

## 3. Core Principles
List **non-negotiable rules**.

- Raw data must never be manually edited.
- All derived data must be reproducible from raw data.
- The pipeline must be deterministic.
- New Pokémon generations should require no code changes.
- Excel is a consumer of data, never the source of truth.

---

## 4. Data Philosophy
Explain the **mental model**.

- Raw API data is treated as immutable source material.
- Transformations are applied only during export.
- One Pokémon form equals one row in the dataset.
- Species-level and form-level data are intentionally separated.

---

## 5. Update Strategy
Explain **how updates work**.

- Raw data is cached locally to avoid redundant API calls.
- Only missing or stale records are refetched.
- Export can be rerun at any time without refetching.
- Changes in formatting or naming require no refetch.

---

## 6. Output Contract
Describe the **final artifact**.

- Primary output: `pokedata.xlsx`
- Contains structured sheets for:
  - Pokémon (forms)
  - Moves
  - Items
  - Abilities
  - Natures
  - Type effectiveness
- Output is designed to be consumed by Excel-based tools.

---

## 7. Long-Term Vision
Describe **where this can grow**, without committing to it.

- Support additional version groups or generations.
- Add alternative export formats (CSV, JSON).
- Enable future migration to a web or application frontend.
- Maintain compatibility with existing Excel tooling.

---

## 8. Success Criteria
Define **how you know it works**.

- A full dataset can be generated with a single command.
- Re-running the pipeline produces identical results.
- Adding new Pokémon or forms requires no manual intervention.
- Excel tooling continues to function without changes after updates.
