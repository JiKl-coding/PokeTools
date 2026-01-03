# Export Prompt – Pokédex Data Pipeline

## Requirements & Source of Truth

This prompt MUST be implemented strictly according to:

- `spec/north-star.md`
- `spec/data-contract.md`
- `spec/export.spec.md`

Rules:
- Sheet structure, column order and formats are binding
- No derived logic allowed
- Export must not modify transformed data

---

## Index
1. Export MVP
2. Export Extended
3. Export Production

---

## 1. Export MVP

### Scope
Export:
- Pokemon sheet
- Meta sheet

---

## 2. Export Extended

### Scope
Add:
- Learnsets
- Moves
- Items
- Abilities
- Natures
- Evolutions
- TypeChart

---

## 3. Export Production

### Scope
Export full workbook defined in specs.

---

## Non-Goals
- No fetching
- No transforming
