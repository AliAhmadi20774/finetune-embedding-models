from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.data import export_duplicate_question_extras


def main() -> None:
    parser = argparse.ArgumentParser(description="Export additional occurrences of normalized duplicate questions.")
    parser.add_argument("--input", type=Path, default=Path("porseman_clean.csv"))
    parser.add_argument("--output", type=Path, default=Path("data/duplicates/duplicate_question_extras.csv"))
    args = parser.parse_args()
    report = export_duplicate_question_extras(args.input, args.output)
    print(json.dumps({"output": str(args.output), **report}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
