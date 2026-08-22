"""Write a JSONL copy containing only the first instance of each query."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def deduplicate(input_path: Path, output_path: Path) -> tuple[int, int]:
    """Return (total_records, removed_records)."""
    seen_queries: set[str] = set()
    total = 0
    removed = 0

    with input_path.open("r", encoding="utf-8") as source, output_path.open(
        "w", encoding="utf-8", newline="\n"
    ) as destination:
        for line_number, line in enumerate(source, start=1):
            if not line.strip():
                continue

            total += 1
            try:
                record = json.loads(line)
                query = record["query"]
            except (json.JSONDecodeError, KeyError) as error:
                raise ValueError(f"Invalid record at line {line_number}: {error}") from error

            if not isinstance(query, str):
                raise ValueError(f"The query at line {line_number} must be a string.")

            if query in seen_queries:
                removed += 1
                continue

            seen_queries.add(query)
            destination.write(line.rstrip("\r\n") + "\n")

    return total, removed


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Remove duplicate JSONL records based on their exact query value."
    )
    parser.add_argument(
        "input",
        nargs="?",
        type=Path,
        default=Path("data/hard_negatives_all.jsonl"),
        help="Source JSONL file.",
    )
    parser.add_argument(
        "output",
        nargs="?",
        type=Path,
        default=Path("data/hard_negatives_deduplicated.jsonl"),
        help="Destination JSONL file.",
    )
    args = parser.parse_args()

    if not args.input.is_file():
        parser.error(f"Input file not found: {args.input}")
    if args.input.resolve() == args.output.resolve():
        parser.error("Output path must be different from the input path.")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    total, removed = deduplicate(args.input, args.output)
    print(f"Processed {total} records; removed {removed} duplicates.")
    print(f"Saved {total - removed} unique-query records to {args.output}")


if __name__ == "__main__":
    main()
