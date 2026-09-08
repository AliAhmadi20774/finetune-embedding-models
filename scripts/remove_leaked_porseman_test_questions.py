"""Create a leakage-filtered Porseman test CSV from a similarity report."""

from __future__ import annotations

import csv
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEST_PATH = PROJECT_ROOT / "data" / "processed" / "porseman_test.csv"
REPORT_PATH = (
    PROJECT_ROOT
    / "outputs"
    / "porseman_train_test_similarity"
    / "best_train_match_for_each_test_question.csv"
)
OUTPUT_PATH = PROJECT_ROOT / "data" / "processed" / "porseman_test2.csv"


def read_csv(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        return list(reader), reader.fieldnames or []


def main() -> None:
    test_rows, fieldnames = read_csv(TEST_PATH)
    report_rows, _ = read_csv(REPORT_PATH)
    leaked_ids = {
        row["test_id"]
        for row in report_rows
        if row.get("meets_threshold", "").strip().lower() == "true"
    }
    test_ids = {row["id"] for row in test_rows}
    unknown_ids = leaked_ids - test_ids
    if unknown_ids:
        raise ValueError(f"Report contains IDs absent from test CSV: {sorted(unknown_ids)[:5]}")

    kept_rows = [row for row in test_rows if row["id"] not in leaked_ids]
    if len(test_rows) - len(kept_rows) != len(leaked_ids):
        raise RuntimeError("Unexpected number of removed rows.")

    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(kept_rows)

    print(f"Original test rows: {len(test_rows):,}")
    print(f"Removed leaked rows: {len(leaked_ids):,}")
    print(f"Kept rows: {len(kept_rows):,}")
    print(f"Written: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
