
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

args = SimpleNamespace(
    filename="工程.csv",  # Default CSV filename
    limit=3,
    pretty=True
)


def available_csvs() -> list[str]:
    """Return a sorted list of available CSV files in the public directory."""
    return sorted(p.name for p in PUBLIC_DIR.glob("*.csv"))



def main(argv: Sequence[str] | None = None) -> int:
    """Main CLI entry point. Reads, filters, and exports CSV data."""
    try:
        files = available_csvs()
        if not files:
            logging.error(f"No CSV files found under {PUBLIC_DIR}.")
            return 1
        logging.info(f"Target filename: {args.filename}")
        target = args.filename or files[0]
        if target not in files:
            logging.error(f"{target} not found in public directory.")
            logging.info("Available files:")
            for name in files:
                logging.info(f"  - {name}")
            return 1

        try:
            rows = read_public_csv(target)
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
            return 1

        filtered = filter_by_binding(rows)
        logging.info(f"Loaded {len(rows)} rows from {target}.")
        logging.info(f"{len(filtered)} filtered row(s)")

        # Save filtered results to a JSON file
        save_path = "./runMacro/target_macro_input.json"
        try:
            with open(save_path, "w", encoding="utf-8") as outjson:
                json.dump(filtered, outjson, ensure_ascii=False, indent=2)
            logging.info(f"Filtered data saved to {save_path}")
        except Exception as e:
            logging.error(f"Error saving JSON: {e}")
            return 1

        limit = max(args.limit, 0)
        if limit and filtered:
            logging.info(f"Showing first {limit} filtered row(s):")
            sample = filtered[:limit]
            if args.pretty:
                print(json.dumps(sample, ensure_ascii=False, indent=2))
            else:
                for idx, row in enumerate(sample, start=1):
                    print(f"[{idx}] {row}")

        return 0
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        return 1



if __name__ == "__main__":
    raise SystemExit(main())
