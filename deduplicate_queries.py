"""Clean hard-negative records and remove duplicate queries from a JSONL file."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def collect_positive_documents(input_path: Path) -> list[str]:
    """Collect distinct positive documents for deterministic negative replacement."""
    positives: dict[str, None] = {}

    with input_path.open("r", encoding="utf-8") as source:
        for line_number, line in enumerate(source, start=1):
            if not line.strip():
                continue
            try:
                record = json.loads(line)
                positive = record["positive_document"]
            except (json.JSONDecodeError, KeyError) as error:
                raise ValueError(f"Invalid record at line {line_number}: {error}") from error

            if not isinstance(positive, str) or not positive.strip():
                raise ValueError(f"Invalid positive_document at line {line_number}.")
            positives[positive.strip()] = None

    return list(positives)


def clean_negatives(
    record: dict[str, Any],
    positive_pool: list[str],
    negative_count: int,
    start_index: int,
) -> tuple[list[dict[str, Any]], int, int, int]:
    """Return cleaned negatives and cleanup counters for one record."""
    positive = record["positive_document"]
    query = record["query"]
    hard_negatives = record["hard_negatives"]
    cleaned: list[dict[str, Any]] = []
    seen: set[str] = set()
    duplicate_count = 0
    false_negative_count = 0
    truncated_count = 0

    if not isinstance(hard_negatives, list):
        raise ValueError("hard_negatives must be a list.")

    for item in hard_negatives:
        if not isinstance(item, dict) or not isinstance(item.get("text"), str):
            raise ValueError("Every hard negative must be an object with a text string.")

        text = item["text"].strip()
        if not text:
            raise ValueError("Hard-negative text cannot be empty.")
        if text == positive:
            false_negative_count += 1
            continue
        if text in seen:
            duplicate_count += 1
            continue
        if len(cleaned) >= negative_count:
            truncated_count += 1
            continue

        cleaned_item = dict(item)
        cleaned_item["text"] = text
        cleaned.append(cleaned_item)
        seen.add(text)

    pool_index = start_index
    attempts = 0
    max_attempts = len(positive_pool) * 2
    while len(cleaned) < negative_count and attempts < max_attempts:
        candidate = positive_pool[pool_index % len(positive_pool)]
        pool_index += 1
        attempts += 1
        if candidate == positive or candidate == query or candidate in seen:
            continue
        cleaned.append(
            {
                "type": "random_negative",
                "hardness": "random",
                "violated_requirement": "Added during dataset cleaning.",
                "text": candidate,
            }
        )
        seen.add(candidate)

    if len(cleaned) != negative_count:
        raise ValueError("Not enough distinct documents to fill the negative list.")

    for item_id, item in enumerate(cleaned, start=1):
        item["id"] = item_id

    return cleaned, duplicate_count, false_negative_count, truncated_count


def clean_and_deduplicate_dataset(
    input_path: Path, output_path: Path, negative_count: int = 7
) -> dict[str, int]:
    """Clean hard negatives, remove repeated queries, and return processing stats."""
    positive_pool = collect_positive_documents(input_path)
    if len(positive_pool) < 2:
        raise ValueError("At least two distinct positive documents are required.")

    seen_queries: set[str] = set()
    stats = {
        "total_records": 0,
        "duplicate_queries": 0,
        "duplicate_negatives": 0,
        "false_negatives": 0,
        "truncated_negatives": 0,
        "saved_records": 0,
    }

    with input_path.open("r", encoding="utf-8") as source, output_path.open(
        "w", encoding="utf-8", newline="\n"
    ) as destination:
        for line_number, line in enumerate(source, start=1):
            if not line.strip():
                continue

            stats["total_records"] += 1
            try:
                record = json.loads(line)
                query = record["query"]
                positive = record["positive_document"]
            except (json.JSONDecodeError, KeyError) as error:
                raise ValueError(f"Invalid record at line {line_number}: {error}") from error

            if not isinstance(query, str) or not query.strip():
                raise ValueError(f"Invalid query at line {line_number}.")
            if not isinstance(positive, str) or not positive.strip():
                raise ValueError(f"Invalid positive_document at line {line_number}.")

            query = query.strip()
            positive = positive.strip()

            if query in seen_queries:
                stats["duplicate_queries"] += 1
                continue

            seen_queries.add(query)
            record["query"] = query
            record["positive_document"] = positive
            (
                record["hard_negatives"],
                duplicate_count,
                false_negative_count,
                truncated_count,
            ) = clean_negatives(
                record,
                positive_pool,
                negative_count,
                start_index=line_number - 1,
            )
            stats["duplicate_negatives"] += duplicate_count
            stats["false_negatives"] += false_negative_count
            stats["truncated_negatives"] += truncated_count
            stats["saved_records"] += 1
            destination.write(json.dumps(record, ensure_ascii=False) + "\n")

    return stats


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean hard negatives and remove records with repeated queries."
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
    parser.add_argument("--negative-count", type=int, default=7)
    args = parser.parse_args()

    if not args.input.is_file():
        parser.error(f"Input file not found: {args.input}")
    if args.input.resolve() == args.output.resolve():
        parser.error("Output path must be different from the input path.")
    if args.negative_count < 1:
        parser.error("--negative-count must be at least 1.")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    stats = clean_and_deduplicate_dataset(
        args.input, args.output, negative_count=args.negative_count
    )
    print(f"Processed {stats['total_records']} records.")
    print(f"Removed {stats['duplicate_queries']} duplicate queries.")
    print(f"Removed {stats['duplicate_negatives']} duplicate negatives.")
    print(f"Removed {stats['false_negatives']} positives mislabeled as negatives.")
    print(f"Removed {stats['truncated_negatives']} excess negatives.")
    print(f"Saved {stats['saved_records']} clean records to {args.output}")


if __name__ == "__main__":
    main()
