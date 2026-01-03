# Fetch Prompt – Pokédex Data Pipeline

## Requirements & Source of Truth

This prompt MUST be implemented strictly according to the following documents:

- `spec/north-star.md`
- `spec/data-contract.md`
- `spec/fetch.spec.md`

Rules:
- Raw entities MUST conform exactly to `data-contract.md`
- File structure, cache behavior, TTL rules and keys are binding
- In case of ambiguity, `fetch.spec.md` and `data-contract.md` take precedence
- No derived logic, transformations, or exports are allowed

---

## Index
1. MVP Fetch (Initial Test)
2. Full Fetch (Sample / Test Mode)
3. Production Fetch (Complete Dataset)

---

## 1. MVP Fetch – Initial Test

### Purpose
Establish the first working milestone of the Pokédex Data Pipeline.

### Scope
Fetch and cache only:
- **RawPokemon**
- **RawSpecies**

### CLI
```bash
python -m src.main fetch pokemon --limit 20
```

### Acceptance Criteria
- Creates raw Pokémon and species cache files
- Re-run skips cached entries
- `--force` refreshes all

---

## 2. Full Fetch – Sample / Test Mode

### Scope
Fetch and cache:
- RawPokemon
- RawSpecies
- RawType
- RawMove
- RawItem
- RawAbility
- RawNature
- RawEvolutionChain (optional)

### CLI
```bash
python -m src.main fetch all --limit 20
```

### Acceptance Criteria
- All entity folders populated
- Cache reuse works correctly

---

## 3. Production Fetch – Complete Dataset

### Scope
Fetch all entities without limits.

### CLI
```bash
python -m src.main fetch all
```

### Acceptance Criteria
- Full raw dataset available
- Ready for transform stage

---

## Non-Goals
- No transforms
- No exports
