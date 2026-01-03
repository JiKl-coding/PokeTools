"""Reference-model transforms for Pokédex Data Pipeline.

Implements derived models for:
- Move
- Item
- Ability
- Nature
- TypeChart (relations list)

All functions are null-safe and deterministic.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional


def _normalize_text(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    return re.sub(r"\s+", " ", value.replace("\f", " ").replace("\n", " ")).strip()


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


def build_move(*, raw_move: Dict[str, Any], about_language: str) -> Optional[Dict[str, Any]]:
    """Build one derived Move record from RawMove."""
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


def build_item(*, raw_item: Dict[str, Any], about_language: str) -> Optional[Dict[str, Any]]:
    """Build one derived Item record from RawItem."""
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


def build_ability(
    *, raw_ability: Dict[str, Any], about_language: str
) -> Optional[Dict[str, Any]]:
    """Build one derived Ability record from RawAbility."""
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


def build_nature(*, raw_nature: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Build one derived Nature record from RawNature."""
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


def build_type_chart_relations(raw_types: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
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
