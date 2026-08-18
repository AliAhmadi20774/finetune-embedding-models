from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.data import load_jsonl
from bge_pipeline.evaluation import pair_diagnostics, retrieval_metrics
from bge_pipeline.modeling import MODEL_NAME, load_model, write_json


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--split", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--batch-size", type=int, default=1)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()
    rows = load_jsonl(args.split)
    model = load_model(args.model)
    report = {
        "model": args.model,
        "split": str(args.split),
        "retrieval": retrieval_metrics(model, rows, args.batch_size),
        "pair_diagnostics": pair_diagnostics(model, rows, args.batch_size, args.seed),
    }
    write_json(report, args.output)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
