from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.data import load_jsonl
from bge_pipeline.evaluation import pair_diagnostics, retrieval_metrics
from bge_pipeline.modeling import MODEL_NAME, load_model, write_json
from bge_pipeline.statistics import compare_train_test, dataset_statistics


def markdown(report: dict) -> str:
    retrieval = report["baseline_test"]["retrieval"]
    similarity = report["distribution_similarity"]
    lines = [
        "# Baseline BGE-M3 and data report", "",
        f"- Model: `{report['model']}`", f"- MRR@10: **{retrieval['mrr@10']:.6f}**",
        f"- Recall@5: **{retrieval['recall@5']:.6f}**",
        f"- Recall@1 / Recall@10: {retrieval['recall@1']:.6f} / {retrieval['recall@10']:.6f}",
        f"- nDCG@10: {retrieval['ndcg@10']:.6f}",
        f"- Train/test distribution check: **{'PASS' if similarity['passed'] else 'FAIL'}**", "",
        "## Train/test statistics", "",
        "| Feature | Train mean | Test mean | Train p50 | Test p50 | Train p90 | Test p90 |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for feature in ("question_characters", "answer_characters", "question_tokens", "answer_tokens"):
        train = report["data_statistics"]["train"].get(feature)
        test = report["data_statistics"]["test"].get(feature)
        if train and test:
            lines.append(f"| {feature} | {train['mean']} | {test['mean']} | {train['p50']} | {test['p50']} | {train['p90']} | {test['p90']} |")
    lines += ["", "The similarity gate requires mean and p50/p90/p95/p99 differences to be at most 10%.", ""]
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Create one baseline retrieval and train/test statistics report.")
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--train-file", type=Path, default=Path("data/processed/train.jsonl"))
    parser.add_argument("--test-file", type=Path, default=Path("data/processed/test.jsonl"))
    parser.add_argument("--output", type=Path, default=Path("reports/baseline_report.json"))
    parser.add_argument("--batch-size", type=int, default=1)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--with-pair-diagnostics", action="store_true", help="Encodes test pairs a second time; slow for 8K documents.")
    args = parser.parse_args()

    train_rows, test_rows = load_jsonl(args.train_file), load_jsonl(args.test_file)
    model = load_model(args.model)
    train_stats = dataset_statistics(train_rows, model.tokenizer)
    test_stats = dataset_statistics(test_rows, model.tokenizer)
    report = {
        "model": args.model,
        "evaluation_stage": "before_fine_tuning",
        "data_statistics": {"train": train_stats, "test": test_stats},
        "distribution_similarity": compare_train_test(train_stats, test_stats),
        "baseline_test": {
            "retrieval": retrieval_metrics(model, test_rows, args.batch_size),
            "pair_diagnostics": pair_diagnostics(model, test_rows, args.batch_size, args.seed) if args.with_pair_diagnostics else None,
        },
    }
    write_json(report, args.output)
    markdown_path = args.output.with_suffix(".md")
    markdown_path.write_text(markdown(report), encoding="utf-8")
    print(json.dumps(report["baseline_test"]["retrieval"], indent=2))
    print(f"Reports: {args.output}, {markdown_path}")
    if not report["distribution_similarity"]["passed"]:
        raise SystemExit("Train/test statistical similarity gate failed; inspect the report before training")


if __name__ == "__main__":
    main()
