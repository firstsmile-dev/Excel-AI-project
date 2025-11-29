"""Utilities for reading CSV data from the local public registry."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Sequence

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


def read_public_csv(filename: str) -> List[Dict[str, str]]:
    """Return the CSV contents as a list of dictionaries.

    The CSV file must be located inside the local `public` directory.
    """

    csv_path = PUBLIC_DIR / filename
    if not csv_path.is_file():
        raise FileNotFoundError(f"{csv_path} does not exist.")

    for encoding in DEFAULT_ENCODINGS:
        try:
            with csv_path.open("r", encoding=encoding, newline="") as fh:
                return list(csv.DictReader(fh))
        except UnicodeDecodeError:
            continue
        except OSError as exc:  # pragma: no cover
            raise CsvReadError(f"Failed to read {csv_path} ({encoding}): {exc}") from exc

    raise CsvReadError(
        f"Unable to decode {csv_path.name} using any of {DEFAULT_ENCODINGS}."
    )


def filter_by_binding(
    rows: Sequence[Dict[str, str]],
    bindings: Sequence[str] = ("Kindle版", "コミック", "雑誌", "大型本", "単行本", "単行本", "ソフトカバー"),
    column: str = BINDING_COLUMN,
) -> List[Dict[str, str]]:
    """Return only rows whose binding column matches any of the desired values (e.g., Kindle, magazine, large format)."""

    return [row for row in rows if row.get(column) in bindings]


if __name__ == "__main__":
    available_files = sorted(p.name for p in PUBLIC_DIR.glob("*.csv"))
    if not available_files:
        print("No CSV files found in the public directory.")
    else:
        first_file = available_files[0]
        rows = read_public_csv(first_file)
        print(f"Loaded {len(rows)} rows from {first_file}.")
