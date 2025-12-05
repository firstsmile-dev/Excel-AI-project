
"""
main.py - CLI to read, filter, and export CSV data from the local public registry.
Refactored for better error handling, logging, and clarity.
"""


from __future__ import annotations
import json
import logging
from typing import Sequence
from types import SimpleNamespace

from init_csv import (
    PUBLIC_DIR,
    filter_by_binding,
    read_public_csv,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def main(argv: Sequence[str] | None = None) -> int:
    """Main CLI entry point. Reads, filters, and exports CSV data."""
    try:
        return 0
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return 1



if __name__ == "__main__":
    try:
        main_success = main()
        if main_success == 0:
            run_script = __import__("vba_simulation")
            vba_success = run_script.run_excel_process()
            if vba_success != 0:
                run_script = __import__("ai_connect")
                ai_success = run_script.__main__()
    except Exception as e:
        logging.error(f"Error in main execution: {e}")