"""CLI entry point for PokéTools.

Delegates to the fetch module so the project supports
running via `python -m src.main`.
"""
from .fetch import main as fetch_main

if __name__ == "__main__":
    # Delegate to fetch/main to support `python -m src.main`
    fetch_main()
