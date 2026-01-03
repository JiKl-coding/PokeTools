"""Export layer for Pokédex Data Pipeline.

Implements Export MVP:
- Writes `data/export/pokedata.xlsx`
- Exports only the `Pokemon` and `Meta` sheets

Rules:
- Sheet structure and column order are binding per spec/export.spec.md
- No derived logic (read transformed data only)
- Export must not modify transformed data
"""

from __future__ import annotations

import os
from typing import Any, Dict, List, Optional

from openpyxl import Workbook

from .cache.io import ensure_dir, read_json
from .transform import load_config


DERIVED_DIR = "data/derived"
EXPORT_DIR = "data/export"
EXPORT_PATH = os.path.join(EXPORT_DIR, "pokedata.xlsx")


def _read_derived(path: str) -> Dict[str, Any]:
    data = read_json(path)
    if not data:
        raise RuntimeError(f"Missing or invalid derived file: {path}")
    return data


def _write_row(ws: Any, values: List[Any]) -> None:
    ws.append(values)


def _moveset_map(
    learnset_entries: List[Dict[str, Any]],
    version_groups: List[str],
) -> Dict[str, Dict[str, str]]:
    """Build {form_key: {version_group: 'a;b;c'}} mapping for Pokemon sheet."""
    by_form_vg: Dict[str, Dict[str, set[str]]] = {}
    vg_allow = set(version_groups)
    for rec in learnset_entries:
        form_key = rec.get("form_key")
        vg = rec.get("version_group")
        move_key = rec.get("move_key")
        if not isinstance(form_key, str) or not isinstance(vg, str) or \
           not isinstance(move_key, str):
            continue
        if vg not in vg_allow:
            continue
        by_form_vg.setdefault(form_key, {}).setdefault(vg, set()).add(move_key)

    out: Dict[str, Dict[str, str]] = {}
    for form_key, vg_map in by_form_vg.items():
        out[form_key] = {}
        for vg, moves in vg_map.items():
            out[form_key][vg] = ";".join(sorted(moves))
    return out


def run_export_mvp(config_path: str = "config/config.json") -> int:
    """Export MVP: create workbook with Pokemon + Meta sheets."""
    cfg = load_config(config_path)
    version_groups = cfg.get("version_groups", [])
    version_groups = [vg for vg in version_groups if isinstance(vg, str)]

    forms_path = os.path.join(DERIVED_DIR, "pokemon_forms.json")
    meta_path = os.path.join(DERIVED_DIR, "meta.json")

    forms_doc = _read_derived(forms_path)
    meta_doc = _read_derived(meta_path)

    forms = forms_doc.get("pokemon_forms")
    if not isinstance(forms, list):
        raise RuntimeError("Derived pokemon_forms.json missing pokemon_forms list")

    ensure_dir(EXPORT_DIR)

    wb = Workbook()
    # Remove default sheet so we control sheet order.
    default_ws = wb.active
    wb.remove(default_ws)

    # Sheet: Pokemon
    ws_pokemon = wb.create_sheet("Pokemon")

    headers: List[str] = [
        "DEX_ID",
        "FORM_KEY",
        "DISPLAY_NAME",
        "FORM_GROUP",
        "TYPE1",
        "TYPE2",
        "HP",
        "ATK",
        "DEF",
        "SPA",
        "SPD",
        "SPE",
        "TOTAL",
        "ABILITY1",
        "ABILITY2",
        "HIDDEN_ABILITY",
        "HEIGHT_M",
        "WEIGHT_KG",
        "ABOUT",
        "SPRITE",
        "SHINY_SPRITE",
    ]
    for vg in version_groups:
        headers.append(f"MOVESET_{vg}")
    _write_row(ws_pokemon, headers)

    for rec in forms:
        if not isinstance(rec, dict):
            continue
        row: List[Any] = [
            rec.get("dex_id"),
            rec.get("form_key"),
            rec.get("display_name"),
            rec.get("form_group"),
            rec.get("type1"),
            rec.get("type2"),
            rec.get("base_hp"),
            rec.get("base_atk"),
            rec.get("base_def"),
            rec.get("base_spa"),
            rec.get("base_spd"),
            rec.get("base_spe"),
            rec.get("base_total"),
            rec.get("ability1"),
            rec.get("ability2"),
            rec.get("hidden_ability"),
            rec.get("height_m"),
            rec.get("weight_kg"),
            rec.get("about"),
            rec.get("sprite_url"),
            rec.get("shiny_sprite_url"),
        ]
        # MVP: do not compute movesets; emit empty cells.
        row.extend([None] * len(version_groups))
        _write_row(ws_pokemon, row)

    # Sheet: Meta
    ws_meta = wb.create_sheet("Meta")
    _write_row(ws_meta, ["KEY", "VALUE"])

    def meta_val(key: str) -> Optional[Any]:
        return meta_doc.get(key)

    required_meta_rows: List[List[Any]] = [
        ["generated_at", meta_val("generated_at")],
        ["source", meta_val("source")],
        ["pokeapi_base_url", cfg.get("pokeapi_base_url")],
        ["about_language", meta_val("about_language")],
        [
            "version_groups",
            ",".join(meta_val("version_groups") or [])
            if isinstance(meta_val("version_groups"), list)
            else "",
        ],
        ["pipeline_version", meta_val("pipeline_version")],
    ]
    for r in required_meta_rows:
        _write_row(ws_meta, r)

    wb.save(EXPORT_PATH)
    print(f"export mvp: wrote {EXPORT_PATH}")
    return 0


def run_export_extended(config_path: str = "config/config.json") -> int:
    """Export Extended: add Learnsets/Moves/Items/Abilities/Natures/Evolutions/TypeChart."""
    cfg = load_config(config_path)
    version_groups = cfg.get("version_groups", [])
    version_groups = [vg for vg in version_groups if isinstance(vg, str)]

    forms_doc = _read_derived(os.path.join(DERIVED_DIR, "pokemon_forms.json"))
    learnsets_doc = _read_derived(os.path.join(DERIVED_DIR, "learnset_entries.json"))
    moves_doc = _read_derived(os.path.join(DERIVED_DIR, "moves.json"))
    items_doc = _read_derived(os.path.join(DERIVED_DIR, "items.json"))
    abilities_doc = _read_derived(os.path.join(DERIVED_DIR, "abilities.json"))
    natures_doc = _read_derived(os.path.join(DERIVED_DIR, "natures.json"))
    evolutions_doc = _read_derived(os.path.join(DERIVED_DIR, "evolution_edges.json"))
    type_chart_doc = _read_derived(os.path.join(DERIVED_DIR, "type_chart.json"))
    meta_doc = _read_derived(os.path.join(DERIVED_DIR, "meta.json"))

    forms = forms_doc.get("pokemon_forms")
    learnset_entries = learnsets_doc.get("learnset_entries")
    moves = moves_doc.get("moves")
    items = items_doc.get("items")
    abilities = abilities_doc.get("abilities")
    natures = natures_doc.get("natures")
    evolutions = evolutions_doc.get("evolution_edges")
    type_relations = type_chart_doc.get("type_chart_relations")

    if not isinstance(forms, list):
        raise RuntimeError("pokemon_forms.json missing pokemon_forms list")
    if not isinstance(learnset_entries, list):
        raise RuntimeError("learnset_entries.json missing learnset_entries list")
    if not isinstance(moves, list):
        raise RuntimeError("moves.json missing moves list")
    if not isinstance(items, list):
        raise RuntimeError("items.json missing items list")
    if not isinstance(abilities, list):
        raise RuntimeError("abilities.json missing abilities list")
    if not isinstance(natures, list):
        raise RuntimeError("natures.json missing natures list")
    if not isinstance(evolutions, list):
        raise RuntimeError("evolution_edges.json missing evolution_edges list")
    if not isinstance(type_relations, list):
        raise RuntimeError("type_chart.json missing type_chart_relations list")

    ensure_dir(EXPORT_DIR)

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    # Sheet: Pokemon
    ws_pokemon = wb.create_sheet("Pokemon")
    pokemon_headers: List[str] = [
        "DEX_ID",
        "FORM_KEY",
        "DISPLAY_NAME",
        "FORM_GROUP",
        "TYPE1",
        "TYPE2",
        "HP",
        "ATK",
        "DEF",
        "SPA",
        "SPD",
        "SPE",
        "TOTAL",
        "ABILITY1",
        "ABILITY2",
        "HIDDEN_ABILITY",
        "HEIGHT_M",
        "WEIGHT_KG",
        "ABOUT",
        "SPRITE",
        "SHINY_SPRITE",
    ]
    for vg in version_groups:
        pokemon_headers.append(f"MOVESET_{vg}")
    _write_row(ws_pokemon, pokemon_headers)

    movesets = _moveset_map(learnset_entries, version_groups)

    # Row ordering: DEX_ID, FORM_KEY
    forms_sorted = [rec for rec in forms if isinstance(rec, dict)]
    forms_sorted.sort(key=lambda r: (r.get("dex_id"), r.get("form_key")))
    for rec in forms_sorted:
        form_key = rec.get("form_key")
        row: List[Any] = [
            rec.get("dex_id"),
            form_key,
            rec.get("display_name"),
            rec.get("form_group"),
            rec.get("type1"),
            rec.get("type2"),
            rec.get("base_hp"),
            rec.get("base_atk"),
            rec.get("base_def"),
            rec.get("base_spa"),
            rec.get("base_spd"),
            rec.get("base_spe"),
            rec.get("base_total"),
            rec.get("ability1"),
            rec.get("ability2"),
            rec.get("hidden_ability"),
            rec.get("height_m"),
            rec.get("weight_kg"),
            rec.get("about"),
            rec.get("sprite_url"),
            rec.get("shiny_sprite_url"),
        ]
        if isinstance(form_key, str):
            for vg in version_groups:
                row.append(movesets.get(form_key, {}).get(vg))
        else:
            row.extend([None] * len(version_groups))
        _write_row(ws_pokemon, row)

    # Sheet: Learnsets
    ws_ls = wb.create_sheet("Learnsets")
    _write_row(ws_ls, ["FORM_KEY", "VERSION_GROUP", "MOVE_KEY", "METHOD", "LEVEL"])
    ls_sorted = [rec for rec in learnset_entries if isinstance(rec, dict)]
    ls_sorted.sort(
        key=lambda r: (
            r.get("form_key"),
            r.get("version_group"),
            r.get("move_key"),
            r.get("method"),
            r.get("level"),
        )
    )
    for rec in ls_sorted:
        _write_row(
            ws_ls,
            [
                rec.get("form_key"),
                rec.get("version_group"),
                rec.get("move_key"),
                rec.get("method"),
                rec.get("level"),
            ],
        )

    # Sheet: Moves
    ws_moves = wb.create_sheet("Moves")
    _write_row(
        ws_moves,
        [
            "MOVE_KEY",
            "TYPE",
            "CATEGORY",
            "POWER",
            "ACCURACY",
            "PP",
            "PRIORITY",
            "EFFECT_SHORT",
        ],
    )
    moves_sorted = [rec for rec in moves if isinstance(rec, dict)]
    moves_sorted.sort(key=lambda r: r.get("move_key"))
    for rec in moves_sorted:
        _write_row(
            ws_moves,
            [
                rec.get("move_key"),
                rec.get("type"),
                rec.get("category"),
                rec.get("power"),
                rec.get("accuracy"),
                rec.get("pp"),
                rec.get("priority"),
                rec.get("effect_short"),
            ],
        )

    # Sheet: Items
    ws_items = wb.create_sheet("Items")
    _write_row(ws_items, ["ITEM_KEY", "CATEGORY", "EFFECT_SHORT"])
    items_sorted = [rec for rec in items if isinstance(rec, dict)]
    items_sorted.sort(key=lambda r: r.get("item_key"))
    for rec in items_sorted:
        _write_row(ws_items, [rec.get("item_key"), rec.get("category"), rec.get("effect_short")])

    # Sheet: Abilities
    ws_abilities = wb.create_sheet("Abilities")
    _write_row(ws_abilities, ["ABILITY_KEY", "EFFECT_SHORT"])
    abilities_sorted = [rec for rec in abilities if isinstance(rec, dict)]
    abilities_sorted.sort(key=lambda r: r.get("ability_key"))
    for rec in abilities_sorted:
        _write_row(ws_abilities, [rec.get("ability_key"), rec.get("effect_short")])

    # Sheet: Natures
    ws_natures = wb.create_sheet("Natures")
    _write_row(ws_natures, ["NATURE_KEY", "INCREASED_STAT", "DECREASED_STAT"])
    natures_sorted = [rec for rec in natures if isinstance(rec, dict)]
    natures_sorted.sort(key=lambda r: r.get("nature_key"))
    for rec in natures_sorted:
        _write_row(
            ws_natures,
            [rec.get("nature_key"), rec.get("increased_stat"), rec.get("decreased_stat")],
        )

    # Sheet: Evolutions
    # Note: data-contract.md defines FROM_DEX_ID/TO_DEX_ID etc; export writes derived values.
    ws_evo = wb.create_sheet("Evolutions")
    _write_row(
        ws_evo,
        [
            "FROM_DEX_ID",
            "TO_DEX_ID",
            "TRIGGER",
            "MIN_LEVEL",
            "ITEM_KEY",
            "TIME_OF_DAY",
            "MIN_HAPPINESS",
            "KNOWN_MOVE_KEY",
            "KNOWN_MOVE_TYPE",
            "LOCATION",
            "GENDER",
            "HELD_ITEM_KEY",
        ],
    )
    evo_sorted = [rec for rec in evolutions if isinstance(rec, dict)]
    evo_sorted.sort(
        key=lambda r: (
            r.get("from_dex_id"),
            r.get("to_dex_id"),
            r.get("trigger"),
            r.get("min_level"),
            r.get("item_key"),
        )
    )
    for rec in evo_sorted:
        _write_row(
            ws_evo,
            [
                rec.get("from_dex_id"),
                rec.get("to_dex_id"),
                rec.get("trigger"),
                rec.get("min_level"),
                rec.get("item_key"),
                rec.get("time_of_day"),
                rec.get("min_happiness"),
                rec.get("known_move_key"),
                rec.get("known_move_type"),
                rec.get("location"),
                rec.get("gender"),
                rec.get("held_item_key"),
            ],
        )

    # Sheet: TypeChart (relations list)
    ws_tc = wb.create_sheet("TypeChart")
    _write_row(ws_tc, ["ATTACKING_TYPE", "DEFENDING_TYPE", "MULTIPLIER"])
    tc_sorted = [rec for rec in type_relations if isinstance(rec, dict)]
    tc_sorted.sort(key=lambda r: (r.get("attacking_type"), r.get("defending_type")))
    for rec in tc_sorted:
        _write_row(
            ws_tc,
            [
                rec.get("attacking_type"),
                rec.get("defending_type"),
                rec.get("multiplier"),
            ],
        )

    # Sheet: Meta
    ws_meta = wb.create_sheet("Meta")
    _write_row(ws_meta, ["KEY", "VALUE"])
    required_meta_rows: List[List[Any]] = [
        ["generated_at", meta_doc.get("generated_at")],
        ["source", meta_doc.get("source")],
        ["pokeapi_base_url", cfg.get("pokeapi_base_url")],
        ["about_language", meta_doc.get("about_language")],
        [
            "version_groups",
            ",".join(meta_doc.get("version_groups") or [])
            if isinstance(meta_doc.get("version_groups"), list)
            else "",
        ],
        ["pipeline_version", meta_doc.get("pipeline_version")],
    ]
    for r in required_meta_rows:
        _write_row(ws_meta, r)

    wb.save(EXPORT_PATH)
    print(f"export extended: wrote {EXPORT_PATH}")
    return 0


def run_export_production(config_path: str = "config/config.json") -> int:
    """Export Production: full workbook, fail fast on missing inputs."""
    # Current implementation exports the full workbook (all sheets) already.
    # Production mode's distinct behavior is to fail fast if required inputs
    # are missing; run_export_extended() already raises on missing derived files.
    rc = run_export_extended(config_path)
    print("export production: ok")
    return rc
