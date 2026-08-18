from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.modeling import MODEL_NAME, inspect_model, load_model, write_json


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--output", type=Path, default=Path("reports/model_inspection.json"))
    args = parser.parse_args()
    model = load_model(args.model)
    report = inspect_model(model)
    write_json(report, args.output)
    print(json.dumps(report, indent=2))
    if report["validation_failures"]:
        raise SystemExit("Model structure did not match the pinned BGE-M3 contract")


if __name__ == "__main__":
    main()
