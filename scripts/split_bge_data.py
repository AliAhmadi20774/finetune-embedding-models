"""Convert a clean dataset to BGE-M3 format and create leakage-safe splits."""

from __future__ import annotations

import argparse
import json
import random
from collections import defaultdict
from pathlib import Path
from typing import Any


SPLIT_NAMES = ("train", "validation", "test")


def load_records(input_path: Path) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []

    with input_path.open("r", encoding="utf-8") as source:
        for line_number, line in enumerate(source, start=1):
            if not line.strip():
                continue

            try:
                record = json.loads(line)
            except json.JSONDecodeError as error:
                raise ValueError(f"Invalid JSON at line {line_number}: {error}") from error

            for field in ("query", "positive_document", "hard_negatives"):
                if field not in record:
                    raise ValueError(f"Missing {field!r} at line {line_number}.")

            if not isinstance(record["query"], str) or not record["query"].strip():
                raise ValueError(f"Invalid query at line {line_number}.")
            if not isinstance(record["positive_document"], str) or not record[
                "positive_document"
            ].strip():
                raise ValueError(f"Invalid positive_document at line {line_number}.")
            if not isinstance(record["hard_negatives"], list):
                raise ValueError(f"Invalid hard_negatives at line {line_number}.")

            try:
                to_bge_example(record)
            except ValueError as error:
                raise ValueError(f"Unclean record at line {line_number}: {error}") from error

            records.append(record)

    if not records:
        raise ValueError("The input file contains no records.")

    return records


def split_without_document_leakage(
    records: list[dict[str, Any]], ratios: tuple[float, float, float], rng: random.Random
) -> dict[str, list[dict[str, Any]]]:
    groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records:
        groups[record["positive_document"]].append(record)

    shuffled_groups = list(groups.values())
    rng.shuffle(shuffled_groups)

    targets = [len(records) * ratio for ratio in ratios]
    counts = [0, 0, 0]
    splits: dict[str, list[dict[str, Any]]] = {name: [] for name in SPLIT_NAMES}

    for group in shuffled_groups:
        split_index = min(
            range(3),
            key=lambda index: counts[index] / targets[index],
        )
        split_name = SPLIT_NAMES[split_index]
        splits[split_name].extend(group)
        counts[split_index] += len(group)

    for records_in_split in splits.values():
        rng.shuffle(records_in_split)

    positive_sets = {
        name: {record["positive_document"] for record in split_records}
        for name, split_records in splits.items()
    }
    assert positive_sets["train"].isdisjoint(positive_sets["validation"])
    assert positive_sets["train"].isdisjoint(positive_sets["test"])
    assert positive_sets["validation"].isdisjoint(positive_sets["test"])

    return splits


def to_bge_example(record: dict[str, Any]) -> dict[str, Any]:
    """Convert one validated clean record to the FlagEmbedding data format."""
    positive = record["positive_document"]
    negatives: list[str] = []
    seen: set[str] = set()

    for item in record["hard_negatives"]:
        if not isinstance(item, dict) or not isinstance(item.get("text"), str):
            raise ValueError("Every hard negative must be an object with a text string.")

        text = item["text"].strip()
        if not text:
            raise ValueError("Hard-negative text cannot be empty.")
        if text == positive:
            raise ValueError(
                "A positive document is mislabeled as a negative. Run the cleaning step first."
            )
        if text in seen:
            raise ValueError(
                "Duplicate negatives found. Run the cleaning step first."
            )

        seen.add(text)
        negatives.append(text)

    if not negatives:
        raise ValueError("At least one negative document is required.")

    return {"query": record["query"], "pos": [positive], "neg": negatives}


def write_splits(
    splits: dict[str, list[dict[str, Any]]],
    output_paths: dict[str, Path],
) -> None:
    for split_name in SPLIT_NAMES:
        output_path = output_paths[split_name]
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with output_path.open("w", encoding="utf-8", newline="\n") as destination:
            for record in splits[split_name]:
                example = to_bge_example(record)
                destination.write(json.dumps(example, ensure_ascii=False) + "\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create leakage-safe train, validation, and test JSONL files for BGE-M3."
    )
    parser.add_argument(
        "input",
        type=Path,
        help="Deduplicated source JSONL file.",
    )
    parser.add_argument("train_output", type=Path, help="Training JSONL output file.")
    parser.add_argument(
        "validation_output", type=Path, help="Validation JSONL output file."
    )
    parser.add_argument("test_output", type=Path, help="Test JSONL output file.")
    parser.add_argument("--train-ratio", type=float, required=True)
    parser.add_argument("--validation-ratio", type=float, required=True)
    parser.add_argument("--test-ratio", type=float, required=True)
    parser.add_argument("--seed", type=int, required=True)
    args = parser.parse_args()

    ratios = (args.train_ratio, args.validation_ratio, args.test_ratio)
    if any(ratio <= 0 for ratio in ratios) or abs(sum(ratios) - 100.0) > 1e-9:
        parser.error("Split ratios must be positive percentages that sum to 100.")
    if not args.input.is_file():
        parser.error(f"Input file not found: {args.input}")

    paths = [
        args.input,
        args.train_output,
        args.validation_output,
        args.test_output,
    ]
    resolved_paths = [path.resolve() for path in paths]
    if len(set(resolved_paths)) != len(resolved_paths):
        parser.error("Input and output paths must all be different.")

    return args


def main() -> None:
    args = parse_args()
    rng = random.Random(args.seed)
    ratios = tuple(
        ratio / 100
        for ratio in (args.train_ratio, args.validation_ratio, args.test_ratio)
    )
    output_paths = {
        "train": args.train_output,
        "validation": args.validation_output,
        "test": args.test_output,
    }

    records = load_records(args.input)
    splits = split_without_document_leakage(records, ratios, rng)
    write_splits(splits, output_paths)

    for split_name in SPLIT_NAMES:
        print(
            f"{split_name}: {len(splits[split_name])} records -> "
            f"{output_paths[split_name]}"
        )


if __name__ == "__main__":
    main()
