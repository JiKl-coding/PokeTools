"""File-based JSON cache helpers.

Provides:
- filesystem-safe key normalization
- atomic JSON writes
- TTL-based staleness checks based on cached `_meta.fetched_at`
"""

import json
import os
import re
import tempfile
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Optional

SAFE_FILENAME_RE = re.compile(r"[^A-Za-z0-9_-]+")


def ensure_dir(path: str) -> None:
    """Create a directory (and parents) if missing."""
    os.makedirs(path, exist_ok=True)


def safe_filename(name: str) -> str:
    """Convert an arbitrary key into a filesystem-safe filename stem.

    Allowed characters: ``a-z``, ``0-9``, ``-``, ``_``.
    """
    name = name.strip().lower()
    # Replace disallowed chars with underscore, collapse repeats
    safe = SAFE_FILENAME_RE.sub("_", name)
    safe = re.sub(r"_+", "_", safe).strip("_")
    return safe or "unnamed"


def _parse_iso(dt_str: str) -> Optional[datetime]:
    """Parse a UTC-ish ISO datetime string into a datetime, or return None."""
    try:
        # Handle both ISO with Z and with offset
        if dt_str.endswith("Z"):
            dt_str = dt_str[:-1] + "+00:00"
        return datetime.fromisoformat(dt_str)
    except (TypeError, ValueError):
        return None


def read_json(path: str) -> Optional[Dict[str, Any]]:
    """Read JSON from disk; return None if missing or invalid."""
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return None


def atomic_write_json(path: str, obj: Dict[str, Any]) -> None:
    """Atomically write a JSON file by writing a temp file then renaming."""
    ensure_dir(os.path.dirname(path))
    fd, tmp_path = tempfile.mkstemp(
        prefix=os.path.basename(path) + ".",
        suffix=".tmp",
        dir=os.path.dirname(path),
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)
    finally:
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass


def file_is_stale(path: str, ttl_days: int) -> bool:
    """Return True when the cached file is missing, invalid, or older than TTL."""
    data = read_json(path)
    if not data or "_meta" not in data:
        return True
    fetched_at = data.get("_meta", {}).get("fetched_at")
    dt = _parse_iso(fetched_at) if isinstance(fetched_at, str) else None
    if not dt:
        return True
    # Ensure timezone-aware UTC
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    now = datetime.now(timezone.utc)
    return (now - dt) > timedelta(days=ttl_days)


def wrap_raw(
    url: str,
    payload: Dict[str, Any],
    etag: Optional[str],
    status: Optional[int],
) -> Dict[str, Any]:
    """Wrap an unmodified API payload with required `_meta` fields."""
    return {
        "_meta": {
            "fetched_at": datetime.now(timezone.utc).isoformat(),
            "url": url,
            **({"etag": etag} if etag else {}),
            **({"status": status} if status is not None else {}),
        },
        "data": payload,
    }
