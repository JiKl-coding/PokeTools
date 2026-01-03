"""Learnset transforms for Pokédex Data Pipeline."""

from __future__ import annotations

from typing import Any, Dict, Iterable, List, Optional


def iter_learnset_entries(
    *,
    raw_pokemon: Dict[str, Any],
    version_groups: List[str],
) -> Iterable[Dict[str, Any]]:
    """Yield LearnsetEntry records from a RawPokemon payload.

    Applies config version group filtering and null-safe parsing.
    """
    pokemon_data = raw_pokemon.get("data") or {}
    if not isinstance(pokemon_data, dict):
        return []
    form_key = pokemon_data.get("name")
    if not isinstance(form_key, str):
        return []

    moves = pokemon_data.get("moves")
    if not isinstance(moves, list):
        return []

    out: List[Dict[str, Any]] = []
    vg_allow = set(version_groups)
    for move in moves:
        if not isinstance(move, dict):
            continue
        move_key = ((move.get("move") or {}).get("name"))
        if not isinstance(move_key, str):
            continue

        details = move.get("version_group_details")
        if not isinstance(details, list):
            continue

        for detail in details:
            if not isinstance(detail, dict):
                continue
            vg = ((detail.get("version_group") or {}).get("name"))
            if not isinstance(vg, str):
                continue
            if vg not in vg_allow:
                continue

            method = ((detail.get("move_learn_method") or {}).get("name"))
            if not isinstance(method, str):
                continue
            method = method.lower()

            level_raw = detail.get("level_learned_at")
            level: Optional[int] = None
            if method == "level-up" and isinstance(level_raw, int) and level_raw > 0:
                level = level_raw

            out.append(
                {
                    "form_key": form_key,
                    "version_group": vg,
                    "move_key": move_key,
                    "method": method,
                    "level": level,
                }
            )

    return out
