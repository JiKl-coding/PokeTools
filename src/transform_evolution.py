"""Evolution transforms for Pokédex Data Pipeline."""

from __future__ import annotations

import os
from typing import Any, Dict, Iterable, List

from .cache.io import read_json


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


def species_name_to_id_map(species_dir: str) -> Dict[str, int]:
    """Map species slug name -> dex id from local RawSpecies cache."""
    mapping: Dict[str, int] = {}
    for path in _iter_json_files(species_dir):
        raw_species = read_json(path) or {}
        data = raw_species.get("data") or {}
        if not isinstance(data, dict):
            continue
        species_id = data.get("id")
        name = data.get("name")
        if isinstance(species_id, int) and isinstance(name, str):
            mapping[name] = species_id
    return mapping


def flatten_evolution_chain(
    *,
    chain_node: Dict[str, Any],
    species_name_to_id: Dict[str, int],
    errors: List[str],
) -> List[Dict[str, Any]]:
    """Flatten a PokéAPI evolution chain node into EvolutionEdge rows."""
    edges: List[Dict[str, Any]] = []

    from_species_name = ((chain_node.get("species") or {}).get("name"))
    from_id = (
        species_name_to_id.get(from_species_name)
        if isinstance(from_species_name, str)
        else None
    )

    evolves_to = chain_node.get("evolves_to")
    if not isinstance(evolves_to, list):
        evolves_to = []

    for child in evolves_to:
        if not isinstance(child, dict):
            continue
        to_species_name = ((child.get("species") or {}).get("name"))
        to_id = (
            species_name_to_id.get(to_species_name)
            if isinstance(to_species_name, str)
            else None
        )

        details = child.get("evolution_details")
        if not isinstance(details, list):
            details = []

        if from_id is None or to_id is None:
            errors.append("Missing species id mapping for evolution edge")
        else:
            for detail in details:
                if not isinstance(detail, dict):
                    continue

                trigger = ((detail.get("trigger") or {}).get("name"))
                if not isinstance(trigger, str):
                    trigger = ""

                item_key = ((detail.get("item") or {}).get("name"))
                held_item_key = ((detail.get("held_item") or {}).get("name"))
                known_move_key = ((detail.get("known_move") or {}).get("name"))
                known_move_type = ((detail.get("known_move_type") or {}).get("name"))
                location = ((detail.get("location") or {}).get("name"))

                min_level = detail.get("min_level")
                min_level = min_level if isinstance(min_level, int) else None

                min_happiness = detail.get("min_happiness")
                min_happiness = (
                    min_happiness if isinstance(min_happiness, int) else None
                )

                time_of_day = detail.get("time_of_day")
                time_of_day = (
                    time_of_day
                    if isinstance(time_of_day, str) and time_of_day
                    else None
                )

                gender = detail.get("gender")
                gender = gender if isinstance(gender, int) else None

                edges.append(
                    {
                        "from_dex_id": from_id,
                        "to_dex_id": to_id,
                        "trigger": trigger,
                        "min_level": min_level,
                        "item_key": item_key if isinstance(item_key, str) else None,
                        "time_of_day": time_of_day,
                        "min_happiness": min_happiness,
                        "known_move_key": (
                            known_move_key
                            if isinstance(known_move_key, str)
                            else None
                        ),
                        "known_move_type": (
                            known_move_type
                            if isinstance(known_move_type, str)
                            else None
                        ),
                        "location": location if isinstance(location, str) else None,
                        "gender": gender,
                        "held_item_key": (
                            held_item_key
                            if isinstance(held_item_key, str)
                            else None
                        ),
                    }
                )

        edges.extend(
            flatten_evolution_chain(
                chain_node=child,
                species_name_to_id=species_name_to_id,
                errors=errors,
            )
        )

    return edges
