"""Naming helpers.

Centralizes deterministic display-name formatting per data-contract.md.
"""

from __future__ import annotations


def slug_titlecase(slug: str) -> str:
    """Convert a PokéAPI slug (kebab-case) to title-cased display name.

    Rules (per data-contract.md):
    - Replace '-' with spaces
    - Title Case each token
    - Single-letter tokens are uppercase
    """
    tokens = slug.replace("-", " ").split()
    out_tokens: list[str] = []
    for token in tokens:
        if len(token) == 1:
            out_tokens.append(token.upper())
        else:
            out_tokens.append(token[:1].upper() + token[1:])
    return " ".join(out_tokens)
