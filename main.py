"""
main.py - CLI to read, filter, and export CSV data from the local public registry.
Refactored for better error handling, logging, and clarity.
"""


from __future__ import annotations
import logging
from typing import Sequence
import importlib

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def main(argv: Sequence[str] | None = None) -> int:
    """Main CLI entry point. Reads, filters, and exports CSV data."""
    try:
        # Step 1: Run vba_simulation.py workflow
        vba_module = importlib.import_module("vba_simulation")
        vba_result = vba_module.run_excel_process()
        if vba_result is not None:
            # Step 2: Run ai_connect.py workflow
            ai_module = importlib.import_module("ai_connect")
            ai_module.__main__()
        return 0
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return 1

if __name__ == "__main__":
    main()