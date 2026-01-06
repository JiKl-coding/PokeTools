# info.md

## config/config.json

The main **project configuration file**.  
Defines **where data comes from**, **how long it is cached**, and **which games are included in the export**.

Used for:
- configuring the **PokeAPI base URL**
- controlling **cache TTL** per data type
- selecting the **language** for descriptive fields (`about_language`)
- defining supported games via `version-groups`
- enabling or disabling **evolution data**
- forcing a **full data refresh** regardless of cache state

`version-groups` act as the **source of truth** for which Pokémon games are processed and exported.

---

## assets.csv

A static **asset registry** that is **included in the final export**.

Contains:
- global UI images (e.g. Pokéball, default item icon)
- Pokémon **type icons**
- move **damage class icons** (Physical / Special / Status)
- block-style variants of attack type icons

This file centralizes all external image URLs so the export can reference assets
without hardcoding links elsewhere in the project.

---

## gamesMap.csv

A **mapping file** that converts `version-groups` identifiers into
**human‑readable game names**.

Used to:
- transform technical slugs (e.g. `scarlet-violet`)
- into user-facing labels (e.g. `Scarlet & Violet`)
- when generating the **GAMES** section in the assets/export output

Fallback behavior:
- if `gamesMap.csv` is missing
- or a specific label is not defined  
→ the raw value from `config/config.json` (`version-groups`) is used instead.
