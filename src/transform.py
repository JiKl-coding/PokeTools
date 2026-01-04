"""Transform layer for Pokédex Data Pipeline.

Implements Transform MVP:
- Reads only local raw JSON files from data/raw/**
- Produces derived PokemonForm + Meta as defined in spec/transform.spec.md
  and spec/data-contract.md
- Performs no network access
"""

from __future__ import annotations

import os
import re
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .cache.io import atomic_write_json, ensure_dir, read_json
from .transform_reference import (
    build_ability,
    build_item,
    build_move,
    build_nature,
    build_type,
    build_type_chart_matrix,
)
from .transform_learnset import iter_learnset_entries
from .transform_evolution import flatten_evolution_chain, species_name_to_id_map
from .naming import slug_titlecase

POKEMON_DIR = "data/raw/pokemon"
SPECIES_DIR = "data/raw/species"
EVOLUTION_CHAIN_DIR = "data/raw/evolution-chain"
TYPE_DIR = "data/raw/type"
MOVE_DIR = "data/raw/move"
ITEM_DIR = "data/raw/item"
ABILITY_DIR = "data/raw/ability"
NATURE_DIR = "data/raw/nature"
DERIVED_DIR = "data/derived"


def load_config(config_path: str) -> Dict[str, Any]:
    """Load JSON configuration from ``config_path``.

    Raises ``RuntimeError`` if the file is missing or invalid.
    """
    data = read_json(config_path)
    if not data:
        raise RuntimeError(f"Missing or invalid config: {config_path}")
    return data


def _species_id_from_url(url: str) -> Optional[int]:
    m = re.search(r"/pokemon-species/(\d+)/", url)
    return int(m.group(1)) if m else None


def _iter_json_files(dir_path: str) -> Iterable[str]:
    if not os.path.isdir(dir_path):
        return []
    entries = [
        os.path.join(dir_path, name)
        for name in os.listdir(dir_path)
        if name.lower().endswith(".json")
    ]
    entries.sort()
    return entries


def _normalize_text(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    # PokéAPI flavor text uses newlines and form feed characters.
    return re.sub(r"\s+", " ", value.replace("\f", " ").replace("\n", " ")).strip()


def _pick_about(species_data: Dict[str, Any], about_language: str) -> Optional[str]:
    entries = species_data.get("flavor_text_entries")
    if not isinstance(entries, list):
        return None

    for entry in entries:
        if not isinstance(entry, dict):
            continue
        lang = ((entry.get("language") or {}).get("name"))
        if lang != about_language:
            continue
        text = _normalize_text(entry.get("flavor_text"))
        if text:
            return text
    return None


def _pick_effect_short(entries: Any, language: str) -> Optional[str]:
    if not isinstance(entries, list):
        return None
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        lang = ((entry.get("language") or {}).get("name"))
        if lang != language:
            continue
        short = entry.get("short_effect")
        if isinstance(short, str) and short.strip():
            return _normalize_text(short)
        effect = entry.get("effect")
        if isinstance(effect, str) and effect.strip():
            return _normalize_text(effect)
    return None


def _as_nonempty_str(value: Any) -> Optional[str]:
    if isinstance(value, str) and value.strip():
        return value
    return None


def _find_first_sprite_url(node: Any, preferred_keys: List[str]) -> Optional[str]:
    """Deterministically find the first non-empty URL under known front-facing sprite keys.

    This intentionally searches nested sprite structures (including versions/animated),
    while keeping ordering stable (preferred keys first, then DFS by sorted dict keys).
    """

    if isinstance(node, dict):
        for key in preferred_keys:
            url = _as_nonempty_str(node.get(key))
            if url:
                return url

        for key in sorted(node.keys()):
            child = node.get(key)
            url = _find_first_sprite_url(child, preferred_keys)
            if url:
                return url
        return None

    if isinstance(node, list):
        for item in node:
            url = _find_first_sprite_url(item, preferred_keys)
            if url:
                return url
        return None

    return None


def _pick_sprites(pokemon_data: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    sprites = pokemon_data.get("sprites") or {}
    if not isinstance(sprites, dict):
        return None, None

    other = sprites.get("other") or {}
    official = (other.get("official-artwork") or {}) if isinstance(other, dict) else {}

    sprite_url = _as_nonempty_str(official.get("front_default")) or _as_nonempty_str(
        sprites.get("front_default")
    )
    shiny_url = _as_nonempty_str(official.get("front_shiny")) or _as_nonempty_str(
        sprites.get("front_shiny")
    )

    # If all preferred sprite sources are null, fall back to any available front-facing sprite
    # (including animated or generation-specific sprites), prioritizing visibility.
    if sprite_url is None:
        sprite_url = _find_first_sprite_url(sprites, ["front_default", "front_female"])
    if shiny_url is None:
        shiny_url = _find_first_sprite_url(sprites, ["front_shiny", "front_shiny_female"])

    return (
        sprite_url,
        shiny_url,
    )


def _classify_form_group(form_key: str) -> str:
    key = (form_key or "").lower()

    if "-mega" in key:
        return "Mega"
    if key.endswith("-gmax"):
        return "Gigantamax"

    # Common regional markers in PokéAPI form keys.
    regional_markers = ("-alola", "-galar", "-hisui", "-paldea")
    if any(marker in key for marker in regional_markers):
        return "Regional"

    if "-" in key:
        return "Other"
    return "Standard"


def _extract_types(pokemon_data: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    types = pokemon_data.get("types")
    if not isinstance(types, list) or not types:
        return None, None

    parsed: List[Tuple[int, str]] = []
    for t in types:
        if not isinstance(t, dict):
            continue
        slot = t.get("slot")
        type_name = ((t.get("type") or {}).get("name"))
        if isinstance(slot, int) and isinstance(type_name, str):
            parsed.append((slot, type_name))
    parsed.sort(key=lambda x: x[0])
    if not parsed or parsed[0][0] != 1:
        return None, None

    type1 = parsed[0][1]
    type2 = parsed[1][1] if len(parsed) > 1 and parsed[1][0] == 2 else None
    return type1, type2


def _extract_stats(pokemon_data: Dict[str, Any]) -> Dict[str, int]:
    stats = pokemon_data.get("stats")
    if not isinstance(stats, list):
        stats = []

    by_name: Dict[str, int] = {}
    for s in stats:
        if not isinstance(s, dict):
            continue
        base = s.get("base_stat")
        name = ((s.get("stat") or {}).get("name"))
        if isinstance(base, int) and isinstance(name, str):
            by_name[name] = base

    hp = by_name.get("hp", 0)
    atk = by_name.get("attack", 0)
    deff = by_name.get("defense", 0)
    spa = by_name.get("special-attack", 0)
    spd = by_name.get("special-defense", 0)
    spe = by_name.get("speed", 0)

    total = hp + atk + deff + spa + spd + spe
    return {
        "base_hp": hp,
        "base_atk": atk,
        "base_def": deff,
        "base_spa": spa,
        "base_spd": spd,
        "base_spe": spe,
        "base_total": total,
    }


def _extract_abilities(
    pokemon_data: Dict[str, Any],
) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    abilities = pokemon_data.get("abilities")
    if not isinstance(abilities, list):
        return None, None, None

    parsed: List[Tuple[int, bool, str]] = []
    for a in abilities:
        if not isinstance(a, dict):
            continue
        slot = a.get("slot")
        is_hidden = a.get("is_hidden")
        name = ((a.get("ability") or {}).get("name"))
        if (
            isinstance(slot, int)
            and isinstance(is_hidden, bool)
            and isinstance(name, str)
        ):
            parsed.append((slot, is_hidden, name))
    parsed.sort(key=lambda x: x[0])

    non_hidden = [name for _, hidden, name in parsed if not hidden]
    hidden = [name for _, is_hidden, name in parsed if is_hidden]

    ability1 = non_hidden[0] if len(non_hidden) >= 1 else None
    ability2 = non_hidden[1] if len(non_hidden) >= 2 else None
    hidden_ability = hidden[0] if hidden else None
    return ability1, ability2, hidden_ability


def build_pokemon_form(
    *,
    raw_pokemon: Dict[str, Any],
    raw_species: Dict[str, Any],
    about_language: str,
) -> Optional[Dict[str, Any]]:
    """Build one PokemonForm derived record.

    Returns None when required inputs are missing and the record should be
    skipped.
    """
    pokemon_data = raw_pokemon.get("data") or {}
    species_data = raw_species.get("data") or {}
    if not isinstance(pokemon_data, dict) or not isinstance(species_data, dict):
        return None

    form_key = pokemon_data.get("name")
    dex_id = species_data.get("id")
    if not isinstance(form_key, str) or not isinstance(dex_id, int):
        return None

    type1, type2 = _extract_types(pokemon_data)
    if not type1:
        return None

    stats = _extract_stats(pokemon_data)
    ability1, ability2, hidden_ability = _extract_abilities(pokemon_data)
    sprite_url, shiny_sprite_url = _pick_sprites(pokemon_data)
    about = _pick_about(species_data, about_language)

    height_dm = pokemon_data.get("height")
    weight_hg = pokemon_data.get("weight")
    height_m = (height_dm / 10) if isinstance(height_dm, (int, float)) else None
    weight_kg = (weight_hg / 10) if isinstance(weight_hg, (int, float)) else None

    # Field order is binding per data-contract.md
    return {
        "dex_id": dex_id,
        "form_key": form_key,
        "display_name": slug_titlecase(form_key),
        "form_group": _classify_form_group(form_key),
        "type1": type1,
        "type2": type2,
        "base_hp": stats["base_hp"],
        "base_atk": stats["base_atk"],
        "base_def": stats["base_def"],
        "base_spa": stats["base_spa"],
        "base_spd": stats["base_spd"],
        "base_spe": stats["base_spe"],
        "base_total": stats["base_total"],
        "ability1": ability1,
        "ability2": ability2,
        "hidden_ability": hidden_ability,
        "about": about,
        "sprite_url": sprite_url,
        "shiny_sprite_url": shiny_sprite_url,
        "height_m": height_m,
        "weight_kg": weight_kg,
    }


def run_transform_extended(config_path: str = "config/config.json") -> int:
    """Run Transform Extended: PokemonForm + LearnsetEntry + EvolutionEdge + Meta."""
    cfg = load_config(config_path)
    about_language = str(cfg.get("about_language", "en"))
    version_groups = cfg.get("version_groups", [])
    version_groups = [vg for vg in version_groups if isinstance(vg, str)]

    errors = 0
    forms: List[Dict[str, Any]] = []
    learnsets: List[Dict[str, Any]] = []

    for path in _iter_json_files(POKEMON_DIR):
        raw_pokemon = read_json(path) or {}
        pokemon_data = raw_pokemon.get("data") or {}
        if not isinstance(pokemon_data, dict):
            errors += 1
            continue

        species_url = ((pokemon_data.get("species") or {}).get("url"))
        species_id = (
            _species_id_from_url(species_url) if isinstance(species_url, str) else None
        )
        if species_id is None:
            errors += 1
            continue

        species_path = os.path.join(SPECIES_DIR, f"{species_id}.json")
        raw_species = read_json(species_path)
        if not raw_species:
            errors += 1
            continue

        form = build_pokemon_form(
            raw_pokemon=raw_pokemon,
            raw_species=raw_species,
            about_language=about_language,
        )
        if form is None:
            errors += 1
            continue
        forms.append(form)

        learnsets.extend(
            list(
                iter_learnset_entries(
                    raw_pokemon=raw_pokemon,
                    version_groups=version_groups,
                )
            )
        )

    forms.sort(key=lambda r: (r["dex_id"], r["form_key"]))

    # Deduplicate learnsets by composite key
    seen_ls: set[Tuple[str, str, str, str, Optional[int]]] = set()
    uniq_ls: List[Dict[str, Any]] = []
    for ls in learnsets:
        key = (
            ls["form_key"],
            ls["version_group"],
            ls["move_key"],
            ls["method"],
            ls["level"],
        )
        if key in seen_ls:
            continue
        seen_ls.add(key)
        uniq_ls.append(ls)

    uniq_ls.sort(
        key=lambda r: (
            r["form_key"],
            r["version_group"],
            r["move_key"],
            r["method"],
            (r["level"] is None, r["level"] or 0),
        )
    )

    # Evolution edges
    species_name_to_id = species_name_to_id_map(SPECIES_DIR)
    evo_errors: List[str] = []
    evo_edges: List[Dict[str, Any]] = []
    for path in _iter_json_files(EVOLUTION_CHAIN_DIR):
        raw_chain = read_json(path) or {}
        data = raw_chain.get("data") or {}
        if not isinstance(data, dict):
            errors += 1
            continue
        chain = data.get("chain")
        if not isinstance(chain, dict):
            errors += 1
            continue
        evo_edges.extend(
            flatten_evolution_chain(
                chain_node=chain,
                species_name_to_id=species_name_to_id,
                errors=evo_errors,
            )
        )

    # Deduplicate evolution edges by primary key composite
    seen_evo: set[Tuple[int, int, str, Optional[int], Optional[str]]] = set()
    uniq_evo: List[Dict[str, Any]] = []
    for e in evo_edges:
        key = (
            e["from_dex_id"],
            e["to_dex_id"],
            e["trigger"],
            e["min_level"],
            e["item_key"],
        )
        if key in seen_evo:
            continue
        seen_evo.add(key)
        uniq_evo.append(e)

    uniq_evo.sort(
        key=lambda r: (
            r["from_dex_id"],
            r["to_dex_id"],
            r["trigger"],
            (r["min_level"] is None, r["min_level"] or 0),
            r["item_key"] or "",
        )
    )

    ensure_dir(DERIVED_DIR)
    atomic_write_json(
        os.path.join(DERIVED_DIR, "pokemon_forms.json"),
        {"pokemon_forms": forms},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "learnset_entries.json"),
        {"learnset_entries": uniq_ls},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "evolution_edges.json"),
        {"evolution_edges": uniq_evo},
    )

    meta = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "pipeline_version": str(cfg.get("pipeline_version", "")),
        "about_language": str(cfg.get("about_language", "")),
        "version_groups": cfg.get("version_groups", []),
        "source": "PokéAPI",
    }
    atomic_write_json(os.path.join(DERIVED_DIR, "meta.json"), meta)

    print(
        "transform extended: "
        f"forms={len(forms)} "
        f"learnsets={len(uniq_ls)} "
        f"evolution_edges={len(uniq_evo)}"
    )
    if evo_errors:
        print(f"transform extended: evolution mapping issues: {len(evo_errors)}")
    if errors:
        print(f"transform extended: skipped/errored records: {errors}")

    return 0 if forms else 1


def _build_move(*, raw_move: Dict[str, Any], about_language: str) -> Optional[Dict[str, Any]]:
    data = raw_move.get("data") or {}
    if not isinstance(data, dict):
        return None

    move_key = data.get("name")
    if not isinstance(move_key, str):
        return None

    type_name = ((data.get("type") or {}).get("name"))
    category = ((data.get("damage_class") or {}).get("name"))
    if not isinstance(type_name, str) or not isinstance(category, str):
        return None

    power = data.get("power")
    power = power if isinstance(power, int) else None
    accuracy = data.get("accuracy")
    accuracy = accuracy if isinstance(accuracy, int) else None
    pp = data.get("pp")
    pp = pp if isinstance(pp, int) else None

    priority = data.get("priority")
    if not isinstance(priority, int):
        return None

    effect_short = _pick_effect_short(data.get("effect_entries"), about_language)

    return {
        "move_key": move_key,
        "type": type_name,
        "category": category,
        "power": power,
        "accuracy": accuracy,
        "pp": pp,
        "priority": priority,
        "effect_short": effect_short,
    }


def _build_item(*, raw_item: Dict[str, Any], about_language: str) -> Optional[Dict[str, Any]]:
    data = raw_item.get("data") or {}
    if not isinstance(data, dict):
        return None

    item_key = data.get("name")
    if not isinstance(item_key, str):
        return None

    category = ((data.get("category") or {}).get("name"))
    category = category if isinstance(category, str) else None
    effect_short = _pick_effect_short(data.get("effect_entries"), about_language)

    return {
        "item_key": item_key,
        "category": category,
        "effect_short": effect_short,
    }


def _build_ability(
    *, raw_ability: Dict[str, Any], about_language: str
) -> Optional[Dict[str, Any]]:
    data = raw_ability.get("data") or {}
    if not isinstance(data, dict):
        return None

    ability_key = data.get("name")
    if not isinstance(ability_key, str):
        return None

    effect_short = _pick_effect_short(data.get("effect_entries"), about_language)

    return {
        "ability_key": ability_key,
        "effect_short": effect_short,
    }


def _build_nature(*, raw_nature: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    data = raw_nature.get("data") or {}
    if not isinstance(data, dict):
        return None

    nature_key = data.get("name")
    if not isinstance(nature_key, str):
        return None

    increased_stat = ((data.get("increased_stat") or {}).get("name"))
    increased_stat = increased_stat if isinstance(increased_stat, str) else None
    decreased_stat = ((data.get("decreased_stat") or {}).get("name"))
    decreased_stat = decreased_stat if isinstance(decreased_stat, str) else None

    return {
        "nature_key": nature_key,
        "increased_stat": increased_stat,
        "decreased_stat": decreased_stat,
    }


def _build_type_chart_relations(raw_types: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Build a complete type chart as a relations list.

    Output rows: (attacking_type, defending_type, multiplier)
    Multipliers: 0, 0.5, 1, 2
    """
    types: Dict[str, Dict[str, Any]] = {}
    for raw in raw_types:
        data = raw.get("data") or {}
        if not isinstance(data, dict):
            continue
        name = data.get("name")
        if isinstance(name, str):
            types[name] = data

    type_names = sorted(types.keys())
    relations: List[Dict[str, Any]] = []

    for atk in type_names:
        dmg_rel = (types[atk].get("damage_relations") or {})
        if not isinstance(dmg_rel, dict):
            dmg_rel = {}

        double_to = {
            (t.get("name"))
            for t in (dmg_rel.get("double_damage_to") or [])
            if isinstance(t, dict) and isinstance(t.get("name"), str)
        }
        half_to = {
            (t.get("name"))
            for t in (dmg_rel.get("half_damage_to") or [])
            if isinstance(t, dict) and isinstance(t.get("name"), str)
        }
        no_to = {
            (t.get("name"))
            for t in (dmg_rel.get("no_damage_to") or [])
            if isinstance(t, dict) and isinstance(t.get("name"), str)
        }

        for dfn in type_names:
            multiplier: float = 1.0
            if dfn in no_to:
                multiplier = 0.0
            elif dfn in double_to:
                multiplier = 2.0
            elif dfn in half_to:
                multiplier = 0.5

            relations.append(
                {
                    "attacking_type": atk,
                    "defending_type": dfn,
                    "multiplier": multiplier,
                }
            )

    return relations


def run_transform_production(config_path: str = "config/config.json") -> int:
    """Run Transform Production: full derived dataset from local raw cache."""
    cfg = load_config(config_path)
    about_language = str(cfg.get("about_language", "en"))
    version_groups = cfg.get("version_groups", [])
    version_groups = [vg for vg in version_groups if isinstance(vg, str)]

    errors = 0

    # Forms + learnsets + evolutions reuse Extended implementation logic.
    forms: List[Dict[str, Any]] = []
    learnsets: List[Dict[str, Any]] = []

    for path in _iter_json_files(POKEMON_DIR):
        raw_pokemon = read_json(path) or {}
        pokemon_data = raw_pokemon.get("data") or {}
        if not isinstance(pokemon_data, dict):
            errors += 1
            continue

        species_url = ((pokemon_data.get("species") or {}).get("url"))
        species_id = (
            _species_id_from_url(species_url) if isinstance(species_url, str) else None
        )
        if species_id is None:
            errors += 1
            continue

        species_path = os.path.join(SPECIES_DIR, f"{species_id}.json")
        raw_species = read_json(species_path)
        if not raw_species:
            errors += 1
            continue

        form = build_pokemon_form(
            raw_pokemon=raw_pokemon,
            raw_species=raw_species,
            about_language=about_language,
        )
        if form is None:
            errors += 1
            continue
        forms.append(form)

        learnsets.extend(
            list(
                iter_learnset_entries(
                    raw_pokemon=raw_pokemon,
                    version_groups=version_groups,
                )
            )
        )

    forms.sort(key=lambda r: (r["dex_id"], r["form_key"]))

    seen_ls: set[Tuple[str, str, str, str, Optional[int]]] = set()
    uniq_ls: List[Dict[str, Any]] = []
    for ls in learnsets:
        key = (
            ls["form_key"],
            ls["version_group"],
            ls["move_key"],
            ls["method"],
            ls["level"],
        )
        if key in seen_ls:
            continue
        seen_ls.add(key)
        uniq_ls.append(ls)

    uniq_ls.sort(
        key=lambda r: (
            r["form_key"],
            r["version_group"],
            r["move_key"],
            r["method"],
            (r["level"] is None, r["level"] or 0),
        )
    )

    species_name_to_id = species_name_to_id_map(SPECIES_DIR)
    evo_errors: List[str] = []
    evo_edges: List[Dict[str, Any]] = []
    for path in _iter_json_files(EVOLUTION_CHAIN_DIR):
        raw_chain = read_json(path) or {}
        data = raw_chain.get("data") or {}
        if not isinstance(data, dict):
            errors += 1
            continue
        chain = data.get("chain")
        if not isinstance(chain, dict):
            errors += 1
            continue
        evo_edges.extend(
            flatten_evolution_chain(
                chain_node=chain,
                species_name_to_id=species_name_to_id,
                errors=evo_errors,
            )
        )

    seen_evo: set[Tuple[int, int, str, Optional[int], Optional[str]]] = set()
    uniq_evo: List[Dict[str, Any]] = []
    for e in evo_edges:
        key = (
            e["from_dex_id"],
            e["to_dex_id"],
            e["trigger"],
            e["min_level"],
            e["item_key"],
        )
        if key in seen_evo:
            continue
        seen_evo.add(key)
        uniq_evo.append(e)

    uniq_evo.sort(
        key=lambda r: (
            r["from_dex_id"],
            r["to_dex_id"],
            r["trigger"],
            (r["min_level"] is None, r["min_level"] or 0),
            r["item_key"] or "",
        )
    )

    # Reference entities
    moves: List[Dict[str, Any]] = []
    for path in _iter_json_files(MOVE_DIR):
        raw = read_json(path) or {}
        rec = build_move(raw_move=raw, about_language=about_language)
        if rec is None:
            errors += 1
            continue
        moves.append(rec)
    moves.sort(key=lambda r: r["move_key"])

    items: List[Dict[str, Any]] = []
    for path in _iter_json_files(ITEM_DIR):
        raw = read_json(path) or {}
        rec = build_item(raw_item=raw, about_language=about_language)
        if rec is None:
            errors += 1
            continue
        items.append(rec)
    items.sort(key=lambda r: r["item_key"])

    abilities: List[Dict[str, Any]] = []
    for path in _iter_json_files(ABILITY_DIR):
        raw = read_json(path) or {}
        rec = build_ability(raw_ability=raw, about_language=about_language)
        if rec is None:
            errors += 1
            continue
        abilities.append(rec)
    abilities.sort(key=lambda r: r["ability_key"])

    natures: List[Dict[str, Any]] = []
    for path in _iter_json_files(NATURE_DIR):
        raw = read_json(path) or {}
        rec = build_nature(raw_nature=raw)
        if rec is None:
            errors += 1
            continue
        natures.append(rec)
    natures.sort(key=lambda r: r["nature_key"])

    raw_types: List[Dict[str, Any]] = []
    for path in _iter_json_files(TYPE_DIR):
        raw = read_json(path) or {}
        if raw:
            raw_types.append(raw)

    types: List[Dict[str, Any]] = []
    for raw in raw_types:
        rec = build_type(raw_type=raw)
        if rec is None:
            errors += 1
            continue
        types.append(rec)
    types.sort(key=lambda r: r["type_key"])

    type_chart = build_type_chart_matrix(raw_types)

    ensure_dir(DERIVED_DIR)
    atomic_write_json(
        os.path.join(DERIVED_DIR, "pokemon_forms.json"),
        {"pokemon_forms": forms},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "learnset_entries.json"),
        {"learnset_entries": uniq_ls},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "evolution_edges.json"),
        {"evolution_edges": uniq_evo},
    )
    atomic_write_json(os.path.join(DERIVED_DIR, "moves.json"), {"moves": moves})
    atomic_write_json(os.path.join(DERIVED_DIR, "items.json"), {"items": items})
    atomic_write_json(
        os.path.join(DERIVED_DIR, "abilities.json"),
        {"abilities": abilities},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "natures.json"),
        {"natures": natures},
    )
    atomic_write_json(
        os.path.join(DERIVED_DIR, "type_chart.json"),
        {"type_keys": type_chart.get("type_keys"), "matrix": type_chart.get("matrix")},
    )
    atomic_write_json(os.path.join(DERIVED_DIR, "types.json"), {"types": types})

    meta = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "pipeline_version": str(cfg.get("pipeline_version", "")),
        "about_language": str(cfg.get("about_language", "")),
        "version_groups": cfg.get("version_groups", []),
        "source": "PokéAPI",
    }
    atomic_write_json(os.path.join(DERIVED_DIR, "meta.json"), meta)

    print(
        "transform production: "
        f"forms={len(forms)} "
        f"learnsets={len(uniq_ls)} "
        f"moves={len(moves)} "
        f"items={len(items)} "
        f"abilities={len(abilities)} "
        f"natures={len(natures)} "
        f"types={len(types)} "
        f"evolution_edges={len(uniq_evo)}"
    )
    if evo_errors:
        print(f"transform production: evolution mapping issues: {len(evo_errors)}")
    if errors:
        print(f"transform production: skipped/errored records: {errors}")

    return 0 if forms else 1


def run_transform_mvp(config_path: str = "config/config.json") -> int:
    """Run Transform MVP: PokemonForm + Meta from local raw cache."""
    cfg = load_config(config_path)
    about_language = str(cfg.get("about_language", "en"))

    errors = 0
    forms: List[Dict[str, Any]] = []

    for path in _iter_json_files(POKEMON_DIR):
        raw_pokemon = read_json(path) or {}
        pokemon_data = raw_pokemon.get("data") or {}
        if not isinstance(pokemon_data, dict):
            errors += 1
            continue

        species_url = ((pokemon_data.get("species") or {}).get("url"))
        species_id = _species_id_from_url(species_url) if isinstance(species_url, str) else None
        if species_id is None:
            errors += 1
            continue

        species_path = os.path.join(SPECIES_DIR, f"{species_id}.json")
        raw_species = read_json(species_path)
        if not raw_species:
            errors += 1
            continue

        form = build_pokemon_form(
            raw_pokemon=raw_pokemon,
            raw_species=raw_species,
            about_language=about_language,
        )
        if form is None:
            errors += 1
            continue
        forms.append(form)

    forms.sort(key=lambda r: (r["dex_id"], r["form_key"]))

    ensure_dir(DERIVED_DIR)
    atomic_write_json(os.path.join(DERIVED_DIR, "pokemon_forms.json"), {"pokemon_forms": forms})

    meta = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "pipeline_version": str(cfg.get("pipeline_version", "")),
        "about_language": str(cfg.get("about_language", "")),
        "version_groups": cfg.get("version_groups", []),
        "source": "PokéAPI",
    }
    atomic_write_json(os.path.join(DERIVED_DIR, "meta.json"), meta)

    print(f"transform mvp: wrote {len(forms)} PokemonForm records")
    if errors:
        print(f"transform mvp: skipped/errored records: {errors}")

    return 0 if forms else 1
