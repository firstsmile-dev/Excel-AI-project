
"""
init_csv.py - Utilities for reading and filtering CSV data from the local public registry.
Refactored for better error handling, logging, and clarity.
"""


from __future__ import annotations
import csv
import logging
from pathlib import Path
from typing import Dict, List, Sequence

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

PUBLIC_DIR = Path(__file__).with_name("public")
BINDING_COLUMN = "装丁、製本"
DEFAULT_ENCODINGS: tuple[str, ...] = (
    "utf-8-sig",
    "utf-8",
    "cp932",  # Windows Japanese
    "shift_jis",
)

class CsvReadError(RuntimeError):
    """Raised when we fail to read the requested CSV file."""
    pass

def read_public_csv(filename: str) -> List[Dict[str, str]]:
    """
    Return the CSV contents as a list of dictionaries.
    The CSV file must be located inside the local `public` directory.
    Handles multiple encodings and logs errors.
    """
    csv_path = PUBLIC_DIR / filename
    if not csv_path.is_file():
        logging.error(f"{csv_path} does not exist.")
        raise FileNotFoundError(f"{csv_path} does not exist.")

    for encoding in DEFAULT_ENCODINGS:
        try:
            with csv_path.open("r", encoding=encoding, newline="") as fh:
                logging.info(f"Reading {csv_path} with encoding {encoding}")
                return list(csv.DictReader(fh))
        except UnicodeDecodeError:
            logging.warning(f"UnicodeDecodeError for {csv_path} with encoding {encoding}")
            continue
        except OSError as exc:
            logging.error(f"Failed to read {csv_path} ({encoding}): {exc}")
            raise CsvReadError(f"Failed to read {csv_path} ({encoding}): {exc}") from exc

    logging.error(f"Unable to decode {csv_path.name} using any of {DEFAULT_ENCODINGS}.")
    raise CsvReadError(
        f"Unable to decode {csv_path.name} using any of {DEFAULT_ENCODINGS}."
    )

def filter_by_binding(
    rows: Sequence[Dict[str, str]],
    bindings: Sequence[str] = ("Kindle版", "コミック", "雑誌", "大型本", "単行本", "単行本", "ソフトカバー"),
    column: str = BINDING_COLUMN,
) -> List[Dict[str, str]]:
    """
    Return only rows whose binding column matches any of the desired values (e.g., Kindle, magazine, large format).
    """
    filtered = [row for row in rows if row.get(column) in bindings]
    logging.info(f"Filtered {len(filtered)} rows by binding type.")
    return filtered

if __name__ == "__main__":
    available_files = sorted(p.name for p in PUBLIC_DIR.glob("*.csv"))
    if not available_files:
        logging.error("No CSV files found in the public directory.")
    else:
        first_file = available_files[0]
        try:
            rows = read_public_csv(first_file)
            logging.info(f"Loaded {len(rows)} rows from {first_file}.")
        except Exception as e:
            logging.error(f"Error reading CSV: {e}")
