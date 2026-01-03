# Transform Prompt – Pokédex Data Pipeline

## Requirements & Source of Truth

This prompt MUST be implemented strictly according to:

- `spec/north-star.md`
- `spec/data-contract.md`
- `spec/transform.spec.md`

Rules:
- All derived models MUST conform exactly to `data-contract.md`
- Field names, types, keys, nullability and ordering are binding
- In case of ambiguity, `data-contract.md` and `transform.spec.md` take precedence
- No inferred or additional logic allowed

---

## Index
1. Transform MVP
2. Transform Extended
3. Transform Production

---

## 1. Transform MVP

### Scope
Produce:
- PokemonForm
- Meta

From:
- RawPokemon
- RawSpecies

### Acceptance Criteria
- One PokemonForm per form
- Correct base stats + TOTAL
- Null-safe handling
- Conforms to data-contract

---

## 2. Transform Extended

### Scope
Add:
- LearnsetEntry
- EvolutionEdge

---

## 3. Transform Production

### Scope
Produce full derived dataset:
- PokemonForm
- LearnsetEntry
- Move
- Item
- Ability
- Nature
- TypeChart
- EvolutionEdge
- Meta

---

## Non-Goals
- No fetching
- No Excel export
