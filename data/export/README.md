# pokedata.xlsx Overview

`pokedata.xlsx` is the canonical Excel export produced by the PokéTools pipeline. Each sheet mirrors the derived domain models defined in `spec/data-contract.md` and is written by `src/export.py` in a deterministic order.

## Sheet Guide

All sheets freeze the header row at Excel row 1; data begins on row 2 across the workbook. When a sheet exposes a `KEY` column (e.g., `FORM_KEY`, `MOVE_KEY`, `ITEM_KEY`, `KEY`), use that column to determine the final populated row. For `Evolutions`, use the `FROM_DEX_ID` column for the same purpose.

| Sheet | Purpose | Key Columns |
| --- | --- | --- |
| Pokemon | One row per Pokémon form. Includes dex id, keys, display name, form grouping, typing, all six base stats, stat total, primary/secondary/hidden abilities, height/weight (metric), flavor text, sprite URLs, and one `MOVESET_<version_group>` column per configured version group (a semicolon-separated list of move keys). | `DEX_ID` (col A), `FORM_KEY` (col B), `DISPLAY_NAME` (col C), `FORM_GROUP` (col D), `TYPE1` (col E), `TYPE2` (col F), `HP`…`SPE` (cols G–L), `TOTAL` (col M), `ABILITY1` (col N), `ABILITY2` (col O), `HIDDEN_ABILITY` (col P), `HEIGHT_M` (col Q), `WEIGHT_KG` (col R), `ABOUT` (col S), `SPRITE` (col T), `SHINY_SPRITE` (col U), `MOVESET_<VG>` (cols V+) |
| Learnsets | Detailed learn records per form/version group/move/method. | `FORM_KEY` (col A), `DISPLAY_NAME` (col B), `VERSION_GROUP` (col C), `MOVE_KEY` (col D), `METHOD` (col E), `LEVEL` (col F) |
| Moves | Reference list of moves and their battle stats. | `MOVE_KEY` (col A), `DISPLAY_NAME` (col B), `TYPE` (col C), `CATEGORY` (col D), `POWER` (col E), `ACCURACY` (col F), `PP` (col G), `PRIORITY` (col H), `EFFECT_SHORT` (col I) |
| Items | Reference list of items with category and effect snippet. | `ITEM_KEY` (col A), `DISPLAY_NAME` (col B), `CATEGORY` (col C), `EFFECT_SHORT` (col D) |
| Abilities | Reference list of abilities. | `ABILITY_KEY` (col A), `DISPLAY_NAME` (col B), `EFFECT_SHORT` (col C) |
| Natures | Reference list of natures. | `NATURE_KEY` (col A), `DISPLAY_NAME` (col B), `INCREASED_STAT` (col C), `DECREASED_STAT` (col D) |
| Evolutions | Flattened species-level evolution edges and their triggers. | `FROM_DEX_ID` (col A), `TO_DEX_ID` (col B), `TRIGGER` (col C), `MIN_LEVEL` (col D), `ITEM_KEY` (col E), `TIME_OF_DAY` (col F), `MIN_HAPPINESS` (col G), `KNOWN_MOVE_KEY` (col H), `KNOWN_MOVE_TYPE` (col I), `LOCATION` (col J), `GENDER` (col K), `HELD_ITEM_KEY` (col L) |
| TypeChart | Matrix view of attacking vs defending types using type display names. Ordering follows `config/typesOrder.json`. | `ATTACKING_TYPE` (col A), defending type columns (cols B+) follow the configured order |
| Types | Reference list of Pokémon types with display names and icon URLs. | `DISPLAY_NAME` (col A), `TYPE_KEY` (col B), `ICON_URL` (col C) |
| Meta | Traceability metadata for the export run. | `KEY` (col A), `VALUE` (col B) |
| Assets | Column blocks copied from `config/assets.csv`, plus a `GAMES` column listing the configured version-group labels (with `All` appended). | Each heading occupies two columns (`<Heading>` name column then address column); the blocks fill columns left-to-right (e.g., first heading uses cols A/B, next uses C/D, etc.), and the terminal `GAMES` list occupies the first free column |
| GAMEVERSIONS | Precomputed availability lists for Excel data validation. Columns are `MOVES_ALL` (col A), `POKEMON_ALL` (col B), `ABILITIES_ALL` (col C), `ITEMS_ALL` (col D), followed by repeating quartets per version group—`MOVES_<vg>`, `POKEMON_<vg>`, `ABILITIES_<vg>`, `ITEMS_<vg>`—continuing across columns E+. Values are display names sorted alphabetically. |

## Update Flow

1. Run `python -m src.main transform production` to regenerate the derived JSONs in `data/derived/`.
2. Run `python -m src.main export production` to rebuild `data/export/pokedata.xlsx` using the latest derived data and configuration.
