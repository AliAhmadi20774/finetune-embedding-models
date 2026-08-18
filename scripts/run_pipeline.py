from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def run(*arguments: str) -> None:
    command = [sys.executable, *arguments]
    print("\n>", " ".join(command), flush=True)
    subprocess.run(command, cwd=ROOT, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the platform-independent single-GPU BGE-M3 pipeline.")
    parser.add_argument("--skip-prepare", action="store_true")
    parser.add_argument("--skip-baseline", action="store_true")
    parser.add_argument("--skip-smoke-test", action="store_true")
    parser.add_argument("--outer-batch-size", type=int, default=2)
    parser.add_argument("--mini-batch-size", type=int, default=1)
    parser.add_argument("--epochs", type=float, default=1.0)
    args = parser.parse_args()

    if not args.skip_prepare:
        run("scripts/prepare_data.py", "--input", "porseman_clean.csv", "--output-dir", "data/processed")
        run("scripts/export_duplicates.py")
    run("scripts/inspect_model.py", "--output", "reports/model_inspection.json")
    if not args.skip_baseline:
        run("scripts/baseline_report.py", "--output", "reports/baseline_report.json", "--batch-size", "1")
    if not args.skip_smoke_test:
        run("scripts/train.py", "--smoke-test", "--outer-batch-size", str(args.outer_batch_size), "--mini-batch-size", str(args.mini_batch_size))
    run("scripts/train.py", "--epochs", str(args.epochs), "--outer-batch-size", str(args.outer_batch_size), "--mini-batch-size", str(args.mini_batch_size))
    run("scripts/evaluate.py", "--model", "outputs/bge-m3-dense/run/final", "--split", "data/processed/test.jsonl", "--output", "reports/final_test.json", "--batch-size", "1")


if __name__ == "__main__":
    main()
