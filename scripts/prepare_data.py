from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.data import add_token_statistics, assert_split_integrity, read_and_split, write_outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare deterministic grouped BGE-M3 train/validation/test data.")
    parser.add_argument("--input", type=Path, default=Path("porseman_clean.csv"))
    parser.add_argument("--output-dir", type=Path, default=Path("data/processed"))
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--model", default="BAAI/bge-m3")
    parser.add_argument("--skip-token-stats", action="store_true", help="Useful only for offline preprocessing tests.")
    args = parser.parse_args()

    rows, report = read_and_split(args.input, args.seed)
    if not args.skip_token_stats:
        add_token_statistics(rows, report, args.model)
    write_outputs(rows, report, args.output_dir)
    assert_split_integrity(args.output_dir)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
