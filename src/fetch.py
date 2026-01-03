"""Fetch CLI and cache layer for Pokédex data.

Handles discovery, fetching, and TTL-aware file-based caching
for raw Pokémon and species payloads from PokéAPI.
"""

import argparse
import re
import sys
import time
from typing import Dict, List, Optional, Tuple

import requests
import pokebase
import pokebase.api as pokebase_api
import pokebase.common as pokebase_common

from .cache.io import (
    atomic_write_json,
    ensure_dir,
    file_is_stale,
    read_json,
    safe_filename,
    wrap_raw,
)
from .transform import run_transform_extended, run_transform_mvp, run_transform_production
from .export import run_export_extended, run_export_mvp, run_export_production

POKEMON_DIR = "data/raw/pokemon"
SPECIES_DIR = "data/raw/species"
EVOLUTION_CHAIN_DIR = "data/raw/evolution-chain"
TYPE_DIR = "data/raw/type"
MOVE_DIR = "data/raw/move"
ITEM_DIR = "data/raw/item"
ABILITY_DIR = "data/raw/ability"
NATURE_DIR = "data/raw/nature"


def _configure_pokebase_base_url(base_url: str) -> None:
    """Configure pokebase to use the configured PokéAPI base URL."""
    pokebase_common.BASE_URL = base_url.rstrip("/")


def _should_retry_http(status_code: Optional[int]) -> bool:
    if status_code is None:
        return True
    if status_code == 429:
        return True
    return 500 <= status_code <= 599


def _extract_status_code(exc: Exception) -> Optional[int]:
    resp = getattr(exc, "response", None)
    return getattr(resp, "status_code", None)


def fetch_resource(
    *,
    base_url: str,
    endpoint: str,
    resource_id: str | int,
    max_retries: int,
    retry_backoff_seconds: float,
    request_delay_seconds: float,
) -> Tuple[str, Optional[str], Optional[int], Dict]:
    """Fetch a detail record via pokebase.

    Returns ``(url, etag, status, payload)``.
    Note: pokebase does not expose response headers/status, so ``etag`` and
    ``status`` may be ``None``.
    """
    _configure_pokebase_base_url(base_url)

    resolved_id: str | int = resource_id
    url: str

    if isinstance(resource_id, str):
        # pokebase validates resource ids as integers at the HTTP layer.
        # Resolve slugs to numeric ids via the convenience wrappers, but keep
        # the Meta URL aligned to the slug-based endpoints.
        if endpoint == "pokemon":
            resolved_id = pokebase.pokemon(resource_id).id_
            url = f"{base_url.rstrip('/')}/pokemon/{resource_id}"
        elif endpoint == "move":
            resolved_id = pokebase.move(resource_id).id_
            url = f"{base_url.rstrip('/')}/move/{resource_id}"
        elif endpoint == "item":
            resolved_id = pokebase.item(resource_id).id_
            url = f"{base_url.rstrip('/')}/item/{resource_id}"
        elif endpoint == "ability":
            resolved_id = pokebase.ability(resource_id).id_
            url = f"{base_url.rstrip('/')}/ability/{resource_id}"
        elif endpoint == "nature":
            resolved_id = pokebase.nature(resource_id).id_
            url = f"{base_url.rstrip('/')}/nature/{resource_id}"
        elif endpoint == "type":
            resolved_id = pokebase.type_(resource_id).id_
            url = f"{base_url.rstrip('/')}/type/{resource_id}"
        else:
            url = pokebase_common.api_url_build(endpoint, resolved_id)
    else:
        url = pokebase_common.api_url_build(endpoint, resolved_id)

    attempt = 0
    while True:
        try:
            payload = pokebase_api.get_data(endpoint, resolved_id, force_lookup=True)
            if request_delay_seconds > 0:
                time.sleep(request_delay_seconds)
            return url, None, None, payload
        except requests.exceptions.HTTPError as exc:
            status_code = _extract_status_code(exc)
            if status_code == 404:
                raise
            if attempt >= max_retries or not _should_retry_http(status_code):
                raise
        except requests.exceptions.RequestException:
            if attempt >= max_retries:
                raise

        attempt += 1
        backoff = retry_backoff_seconds * (2 ** (attempt - 1))
        time.sleep(backoff)


def load_config(config_path: str) -> Dict:
    """Load JSON configuration from ``config_path``.

    Raises ``RuntimeError`` if the file is missing or invalid.
    """
    data = read_json(config_path)
    if not data:
        raise RuntimeError(f"Missing or invalid config: {config_path}")
    return data


def discover_form_keys(base_url: str) -> List[str]:
    """Discover all Pokémon form keys via the listing endpoint.

    Falls back to ``pokebase.APIResourceList`` if the HTTP listing is empty.
    Returns a list of form key names.
    """
    # Prefer direct HTTP for the big list; pokebase can be memory-heavy here
    url = f"{base_url.rstrip('/')}/pokemon?limit=100000&offset=0"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    if not results:
        # Fallback: pokebase APIResourceList
        try:
            arl = pokebase.APIResourceList("pokemon")
            return [item["name"] for item in arl]
        except requests.exceptions.RequestException:
            return []
        except (KeyError, TypeError, ValueError):
            return []
    return [r["name"] for r in results]


def species_id_from_url(url: str) -> Optional[int]:
    """Extract numeric species id from a PokéAPI species URL.

    Returns ``None`` when the URL does not match the expected pattern.
    """
    m = re.search(r"/pokemon-species/(\d+)/", url)
    return int(m.group(1)) if m else None


def fetch_json(url: str) -> Tuple[int, Optional[str], Dict]:
    """Fetch JSON from ``url`` returning ``(status, etag, payload)``.

    This is used only for list/discovery endpoints.
    """
    resp = requests.get(url, timeout=30)
    status = resp.status_code
    etag = resp.headers.get("ETag")
    resp.raise_for_status()
    return status, etag, resp.json()


def ensure_dirs() -> None:
    """Ensure cache directories exist for all raw entities."""
    ensure_dir(POKEMON_DIR)
    ensure_dir(SPECIES_DIR)
    ensure_dir(EVOLUTION_CHAIN_DIR)
    ensure_dir(TYPE_DIR)
    ensure_dir(MOVE_DIR)
    ensure_dir(ITEM_DIR)
    ensure_dir(ABILITY_DIR)
    ensure_dir(NATURE_DIR)


def fetch_pokemon_and_species(
    base_url: str,
    name: str,
    ttl_pokemon: int,
    ttl_species: int,
    force: bool,
    max_retries: int,
    retry_backoff_seconds: float,
    request_delay_seconds: float,
) -> Tuple[str, Optional[int], str]:
    """Fetch a Pokémon by form key and its corresponding species.

    Returns ``(pokemon_status, species_id, species_status)`` where each status is
    one of: ``skipped``, ``fetched_new``, ``refreshed``, ``failed``.
    """
    ensure_dirs()
    safe_name = safe_filename(name)
    pokemon_path = f"{POKEMON_DIR}/{safe_name}.json"

    pokemon_status: str
    species_status: str = "skipped"
    species_id: Optional[int] = None

    try:
        need_fetch_pokemon = force or file_is_stale(pokemon_path, ttl_pokemon)
        if not need_fetch_pokemon:
            pokemon_status = "skipped"
            # Load to derive species id
            existing = read_json(pokemon_path) or {}
            species = (existing.get("data") or {}).get("species") or {}
            species_id = species_id_from_url(species.get("url", ""))
        else:
            url, etag, status, payload = fetch_resource(
                base_url=base_url,
                endpoint="pokemon",
                resource_id=name,
                max_retries=max_retries,
                retry_backoff_seconds=retry_backoff_seconds,
                request_delay_seconds=request_delay_seconds,
            )
            wrapped = wrap_raw(url, payload, etag, status)
            existed = read_json(pokemon_path) is not None
            atomic_write_json(pokemon_path, wrapped)
            pokemon_status = "refreshed" if existed else "fetched_new"
            species_id = species_id_from_url(payload.get("species", {}).get("url", ""))

        # Species handling
        if species_id is not None:
            species_path = f"{SPECIES_DIR}/{species_id}.json"
            need_fetch_species = force or file_is_stale(species_path, ttl_species)
            if not need_fetch_species:
                species_status = "skipped"
            else:
                url, etag, status, payload = fetch_resource(
                    base_url=base_url,
                    endpoint="pokemon-species",
                    resource_id=species_id,
                    max_retries=max_retries,
                    retry_backoff_seconds=retry_backoff_seconds,
                    request_delay_seconds=request_delay_seconds,
                )
                wrapped = wrap_raw(url, payload, etag, status)
                existed = read_json(species_path) is not None
                atomic_write_json(species_path, wrapped)
                species_status = "refreshed" if existed else "fetched_new"
        else:
            species_status = "failed"
    except (requests.exceptions.RequestException, ValueError, OSError):
        return ("failed", None, "failed")

    return (pokemon_status, species_id, species_status)


def run_fetch_pokemon(limit: Optional[int], force: bool, config_path: str) -> int:
    """Run the fetch flow for Pokémon and species.

    Applies TTL rules and optional ``limit`` and ``force`` flags.
    Returns 0 on success, 1 when any failures occur.
    """
    cfg = load_config(config_path)
    base_url = cfg.get("pokeapi_base_url", "https://pokeapi.co/api/v2")
    ttl_pokemon = int(cfg.get("ttl_days", {}).get("pokemon", 7))
    ttl_species = int(cfg.get("ttl_days", {}).get("species", 7))

    max_retries = int(cfg.get("max_retries", 5))
    retry_backoff_seconds = float(cfg.get("retry_backoff_seconds", 1.0))
    request_delay_seconds = float(cfg.get("request_delay_seconds", 0.1))

    keys = discover_form_keys(base_url)
    discovered = len(keys)
    if limit is not None:
        keys = keys[:limit]

    counts = {
        "discovered": discovered,
        "processed": 0,
        "fetched_new": 0,
        "refreshed": 0,
        "skipped": 0,
        "failed": 0,
    }

    print(f"Discovered {discovered} pokemon forms; processing {len(keys)}")

    for name in keys:
        p_status, species_id, s_status = fetch_pokemon_and_species(
            base_url,
            name,
            ttl_pokemon,
            ttl_species,
            force,
            max_retries,
            retry_backoff_seconds,
            request_delay_seconds,
        )
        counts["processed"] += 1
        for st in (p_status, s_status):
            if st in counts:
                counts[st] += 1
            elif st == "failed":
                counts["failed"] += 1

        species_info = (
            f" species={species_id}" if species_id is not None else " species=?"
        )
        msg = (
            f"- {name:20s} -> "
            f"pokemon:{p_status:10s} "
            f"species:{s_status:10s}"
            f"{species_info}"
        )
        print(msg)

    print(
        "Summary: "
        + ", ".join(
            [
                f"discovered={counts['discovered']}",
                f"processed={counts['processed']}",
                f"fetched_new={counts['fetched_new']}",
                f"refreshed={counts['refreshed']}",
                f"skipped={counts['skipped']}",
                f"failed={counts['failed']}",
            ]
        )
    )
    return 0 if counts["failed"] == 0 else 1


def discover_keys(base_url: str, endpoint: str, limit: int) -> List[str]:
    """Discover resource slugs for an endpoint using list transport."""
    url = f"{base_url.rstrip('/')}/{endpoint}?limit={limit}&offset=0"
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    return [r["name"] for r in results]


def _evolution_chain_id_from_url(url: str) -> Optional[int]:
    m = re.search(r"/evolution-chain/(\d+)/", url)
    return int(m.group(1)) if m else None


def fetch_and_cache_entity(
    *,
    base_url: str,
    entity: str,
    endpoint: str,
    key: str | int,
    cache_dir: str,
    ttl_days: int,
    force: bool,
    max_retries: int,
    retry_backoff_seconds: float,
    request_delay_seconds: float,
) -> str:
    """Fetch and cache a single detail record using TTL rules."""
    ensure_dirs()
    file_key = safe_filename(key) if isinstance(key, str) else str(key)
    path = f"{cache_dir}/{file_key}.json"

    need_fetch = force or file_is_stale(path, ttl_days)
    if not need_fetch:
        return "skipped"

    existed = read_json(path) is not None
    try:
        url, etag, status, payload = fetch_resource(
            base_url=base_url,
            endpoint=endpoint,
            resource_id=key,
            max_retries=max_retries,
            retry_backoff_seconds=retry_backoff_seconds,
            request_delay_seconds=request_delay_seconds,
        )
        wrapped = wrap_raw(url, payload, etag, status)
        atomic_write_json(path, wrapped)
        return "refreshed" if existed else "fetched_new"
    except requests.exceptions.HTTPError as exc:
        if _extract_status_code(exc) == 404:
            print(f"- {entity}:{key} -> missing (404)")
            return "failed"
        raise


def run_fetch_reference(limit: Optional[int], force: bool, config_path: str) -> int:
    """Fetch reference entities (type/move/item/ability/nature) with caching."""
    cfg = load_config(config_path)
    base_url = cfg.get("pokeapi_base_url", "https://pokeapi.co/api/v2")
    ttl = cfg.get("ttl_days", {})

    max_retries = int(cfg.get("max_retries", 5))
    retry_backoff_seconds = float(cfg.get("retry_backoff_seconds", 1.0))
    request_delay_seconds = float(cfg.get("request_delay_seconds", 0.1))

    entities = [
        ("type", "type", TYPE_DIR, int(ttl.get("type", 7)), 1000),
        ("move", "move", MOVE_DIR, int(ttl.get("move", 7)), 100000),
        ("item", "item", ITEM_DIR, int(ttl.get("item", 7)), 100000),
        ("ability", "ability", ABILITY_DIR, int(ttl.get("ability", 7)), 100000),
        ("nature", "nature", NATURE_DIR, int(ttl.get("nature", 7)), 1000),
    ]

    total_failed = 0
    for entity, endpoint, cache_dir, ttl_days, discovery_limit in entities:
        keys = discover_keys(base_url, endpoint, discovery_limit)
        discovered = len(keys)
        if limit is not None:
            keys = keys[:limit]

        counts = {
            "discovered": discovered,
            "processed": 0,
            "fetched_new": 0,
            "refreshed": 0,
            "skipped": 0,
            "failed": 0,
        }

        print(f"Discovered {discovered} {entity} keys; processing {len(keys)}")
        for key in keys:
            try:
                status = fetch_and_cache_entity(
                    base_url=base_url,
                    entity=entity,
                    endpoint=endpoint,
                    key=key,
                    cache_dir=cache_dir,
                    ttl_days=ttl_days,
                    force=force,
                    max_retries=max_retries,
                    retry_backoff_seconds=retry_backoff_seconds,
                    request_delay_seconds=request_delay_seconds,
                )
            except (requests.exceptions.RequestException, OSError, ValueError):
                status = "failed"

            counts["processed"] += 1
            if status in counts:
                counts[status] += 1
            else:
                counts["failed"] += 1
            print(f"- {entity}:{key} -> {status}")

        total_failed += counts["failed"]
        print(
            f"Summary {entity}: "
            + ", ".join(
                [
                    f"discovered={counts['discovered']}",
                    f"processed={counts['processed']}",
                    f"fetched_new={counts['fetched_new']}",
                    f"refreshed={counts['refreshed']}",
                    f"skipped={counts['skipped']}",
                    f"failed={counts['failed']}",
                ]
            )
        )

    return 0 if total_failed == 0 else 1


def run_fetch_evolution_chains(
    species_ids: List[int],
    *,
    limit: Optional[int],
    force: bool,
    config_path: str,
) -> int:
    """Fetch evolution chains for the provided species ids (config-gated)."""
    cfg = load_config(config_path)
    base_url = cfg.get("pokeapi_base_url", "https://pokeapi.co/api/v2")
    ttl_days = int(cfg.get("ttl_days", {}).get("evolution_chain", 7))

    max_retries = int(cfg.get("max_retries", 5))
    retry_backoff_seconds = float(cfg.get("retry_backoff_seconds", 1.0))
    request_delay_seconds = float(cfg.get("request_delay_seconds", 0.1))

    include_evolutions = bool(cfg.get("include_evolutions", True))
    if not include_evolutions:
        print("Evolution chains disabled by config (include_evolutions=false)")
        return 0

    chain_ids: List[int] = []
    for species_id in species_ids:
        species_path = f"{SPECIES_DIR}/{species_id}.json"
        data = read_json(species_path) or {}
        evo_url = (data.get("data") or {}).get("evolution_chain", {}).get("url")
        if isinstance(evo_url, str) and evo_url:
            chain_id = _evolution_chain_id_from_url(evo_url)
            if chain_id is not None:
                chain_ids.append(chain_id)

    chain_ids = sorted(set(chain_ids))
    discovered = len(chain_ids)
    if limit is not None:
        chain_ids = chain_ids[:limit]

    print(f"Discovered {discovered} evolution chains; processing {len(chain_ids)}")
    failed = 0
    for chain_id in chain_ids:
        try:
            status = fetch_and_cache_entity(
                base_url=base_url,
                entity="evolution-chain",
                endpoint="evolution-chain",
                key=chain_id,
                cache_dir=EVOLUTION_CHAIN_DIR,
                ttl_days=ttl_days,
                force=force,
                max_retries=max_retries,
                retry_backoff_seconds=retry_backoff_seconds,
                request_delay_seconds=request_delay_seconds,
            )
        except (requests.exceptions.RequestException, OSError, ValueError):
            status = "failed"

        if status == "failed":
            failed += 1
        print(f"- evolution-chain:{chain_id} -> {status}")

    return 0 if failed == 0 else 1


def run_fetch_all(limit: Optional[int], force: bool, config_path: str) -> int:
    """Fetch pokemon+species plus reference entities (and optional evolutions)."""
    cfg = load_config(config_path)
    base_url = cfg.get("pokeapi_base_url", "https://pokeapi.co/api/v2")
    ttl_pokemon = int(cfg.get("ttl_days", {}).get("pokemon", 7))
    ttl_species = int(cfg.get("ttl_days", {}).get("species", 7))

    max_retries = int(cfg.get("max_retries", 5))
    retry_backoff_seconds = float(cfg.get("retry_backoff_seconds", 1.0))
    request_delay_seconds = float(cfg.get("request_delay_seconds", 0.1))

    keys = discover_form_keys(base_url)
    discovered = len(keys)
    if limit is not None:
        keys = keys[:limit]
    print(f"Discovered {discovered} pokemon forms; processing {len(keys)}")

    species_ids: List[int] = []
    failed = 0
    for name in keys:
        p_status, species_id, s_status = fetch_pokemon_and_species(
            base_url,
            name,
            ttl_pokemon,
            ttl_species,
            force,
            max_retries,
            retry_backoff_seconds,
            request_delay_seconds,
        )
        if p_status == "failed" or s_status == "failed":
            failed += 1
        if species_id is not None:
            species_ids.append(species_id)
        species_info = (
            f" species={species_id}" if species_id is not None else " species=?"
        )
        print(
            f"- {name:20s} -> "
            f"pokemon:{p_status:10s} "
            f"species:{s_status:10s}"
            f"{species_info}"
        )

    ref_code = run_fetch_reference(limit, force, config_path)
    evo_code = run_fetch_evolution_chains(
        sorted(set(species_ids)), limit=limit, force=force, config_path=config_path
    )

    return 0 if (failed == 0 and ref_code == 0 and evo_code == 0) else 1


def build_arg_parser() -> argparse.ArgumentParser:
    """Build and return the top-level CLI argument parser."""
    parser = argparse.ArgumentParser(prog="poketools")
    sub = parser.add_subparsers(dest="command")

    fetch = sub.add_parser("fetch")
    fetch_sub = fetch.add_subparsers(dest="entity")

    pokemon = fetch_sub.add_parser("pokemon")
    pokemon.add_argument("--limit", type=int, default=None)
    pokemon.add_argument("--force", action="store_true")

    reference = fetch_sub.add_parser("reference")
    reference.add_argument("--limit", type=int, default=None)
    reference.add_argument("--force", action="store_true")

    all_cmd = fetch_sub.add_parser("all")
    all_cmd.add_argument("--limit", type=int, default=None)
    all_cmd.add_argument("--force", action="store_true")

    transform = sub.add_parser("transform")
    transform_sub = transform.add_subparsers(dest="stage")
    transform_sub.add_parser("mvp")
    transform_sub.add_parser("extended")
    transform_sub.add_parser("production")

    export = sub.add_parser("export")
    export_sub = export.add_subparsers(dest="stage")
    export_sub.add_parser("mvp")
    export_sub.add_parser("extended")
    export_sub.add_parser("production")

    return parser


def main(argv: Optional[List[str]] = None) -> int:
    """CLI entry point for the fetch command."""
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    # Optional default force_refresh from config
    cfg = load_config("config/config.json")
    default_force = bool(cfg.get("force_refresh", False))

    if args.command == "fetch" and args.entity == "pokemon":
        force = bool(args.force) or default_force
        return run_fetch_pokemon(args.limit, force, "config/config.json")

    if args.command == "fetch" and args.entity == "reference":
        force = bool(args.force) or default_force
        return run_fetch_reference(args.limit, force, "config/config.json")

    if args.command == "fetch" and args.entity == "all":
        force = bool(args.force) or default_force
        return run_fetch_all(args.limit, force, "config/config.json")

    if args.command == "transform" and args.stage == "mvp":
        return run_transform_mvp("config/config.json")

    if args.command == "transform" and args.stage == "extended":
        return run_transform_extended("config/config.json")

    if args.command == "transform" and args.stage == "production":
        return run_transform_production("config/config.json")

    if args.command == "export" and args.stage == "mvp":
        return run_export_mvp("config/config.json")

    if args.command == "export" and args.stage == "extended":
        return run_export_extended("config/config.json")

    if args.command == "export" and args.stage == "production":
        return run_export_production("config/config.json")

    parser.print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
