from __future__ import annotations

import math
import statistics
from typing import Any, Callable


PERCENTILES = ("p50", "p90", "p95", "p99")


def _summary(values: list[int]) -> dict[str, float | int]:
    ordered = sorted(values)
    result: dict[str, float | int] = {
        "count": len(values),
        "mean": round(statistics.mean(values), 3),
        "std": round(statistics.pstdev(values), 3),
        "min": ordered[0],
        "max": ordered[-1],
    }
    for name, fraction in (("p50", .5), ("p90", .9), ("p95", .95), ("p99", .99)):
        result[name] = ordered[round((len(ordered) - 1) * fraction)]
    return result


def dataset_statistics(rows: list[dict[str, Any]], tokenizer=None) -> dict[str, Any]:
    question_chars = [len(row["question"]) for row in rows]
    answer_chars = [len(row["content_text"]) for row in rows]
    result: dict[str, Any] = {
        "rows": len(rows),
        "unique_questions": len({row["question_id"] for row in rows}),
        "unique_documents": len({row["document_id"] for row in rows}),
        "question_characters": _summary(question_chars),
        "answer_characters": _summary(answer_chars),
    }
    if tokenizer is not None:
        def token_lengths(field: str, batch_size: int = 128) -> list[int]:
            lengths: list[int] = []
            for start in range(0, len(rows), batch_size):
                texts = [row[field] for row in rows[start:start + batch_size]]
                encoded = tokenizer(texts, add_special_tokens=True, truncation=False)["input_ids"]
                lengths.extend(len(ids) for ids in encoded)
            return lengths

        question_tokens = token_lengths("question")
        answer_tokens = token_lengths("content_text")
        result["question_tokens"] = _summary(question_tokens)
        result["answer_tokens"] = _summary(answer_tokens)
        result["answers_over_8192_tokens"] = sum(value > 8192 for value in answer_tokens)
    return result


def compare_train_test(train: dict[str, Any], test: dict[str, Any], tolerance: float = .10) -> dict[str, Any]:
    """Compare robust distribution landmarks; extremes and std are reported but not gated."""
    checks = []
    for feature in ("question_characters", "answer_characters", "question_tokens", "answer_tokens"):
        if feature not in train or feature not in test:
            continue
        for metric in ("mean", *PERCENTILES):
            left, right = float(train[feature][metric]), float(test[feature][metric])
            relative = abs(right - left) / max(abs(left), 1.0)
            checks.append({
                "feature": feature,
                "metric": metric,
                "train": left,
                "test": right,
                "relative_difference": round(relative, 6),
                "threshold": tolerance,
                "passed": relative <= tolerance,
            })
    return {
        "definition": "Mean and p50/p90/p95/p99 must differ by at most 10%; min/max/std are descriptive only.",
        "passed": all(check["passed"] for check in checks),
        "checks": checks,
    }
