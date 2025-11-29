"""Simple CLI to read CSV data from the local public registry."""

from __future__ import annotations

import json
from typing import Sequence
from types import SimpleNamespace

from init_csv import (
    PUBLIC_DIR,
    filter_by_binding,
    read_public_csv,
)

args = SimpleNamespace(
    filename="工程.csv",
    limit=3,
    pretty=True
)

def available_csvs() -> list[str]:
    return sorted(p.name for p in PUBLIC_DIR.glob("*.csv"))


def main(argv: Sequence[str] | None = None) -> int:
    
    files = available_csvs()
    if not files:
        print(f"No CSV files found under {PUBLIC_DIR}.")
        return 1
    print(args.filename)
    target = args.filename or files[0]
    if target not in files:
        print(f"{target} not found in public directory.")
        print("Available files:")
        for name in files:
            print(f"  - {name}")
        return 1

    rows = read_public_csv(target)
    filtered = filter_by_binding(rows)
    print(f"Loaded {len(rows)} rows from {target}.")
    print(
        f"{len(filtered)} filtered row(s) "
    )
    # INSERT_YOUR_CODE
    # Save filtered results to a JSON file
    # Save the filtered data to a JSON file
    save_path = f"./runMacro/target_macro_input.json"
    with open(save_path, "w", encoding="utf-8") as outjson:
        json.dump(filtered, outjson, ensure_ascii=False, indent=2)
    print(f"Filtered data saved to {save_path}")
    limit = max(args.limit, 0)
    if limit and filtered:
        print(f"Showing first {limit} filtered row(s):")
        sample = filtered[:limit]
        if args.pretty:
            print(json.dumps(sample, ensure_ascii=False, indent=2))
        else:
            for idx, row in enumerate(sample, start=1):
                print(f"[{idx}] {row}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
