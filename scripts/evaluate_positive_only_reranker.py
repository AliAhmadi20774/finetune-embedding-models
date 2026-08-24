"""Evaluate a reranker against the unique positive documents in a test set."""

from __future__ import annotations

import argparse
import json
import platform
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from tqdm import tqdm

from evaluate_positive_only import build_corpus, load_test_records, resolve_output_dir, save_reports


DEFAULT_URL = "http://103.130.147.241:44609/v1/rerank"


def rerank(url: str, model: str, query: str, documents: list[str], timeout: int) -> list[int]:
    """Return document indexes sorted from most to least relevant."""
    payload = json.dumps(
        {"model": model, "query": query, "documents": documents},
        ensure_ascii=False,
    ).encode("utf-8")
    request = Request(url, data=payload, headers={"Content-Type": "application/json"})
    try:
        with urlopen(request, timeout=timeout) as response:
            body = json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Reranker returned HTTP {error.code}: {detail}") from error
    except URLError as error:
        raise RuntimeError(f"Could not reach reranker at {url}: {error.reason}") from error

    results = body.get("results")
    if not isinstance(results, list) or len(results) != len(documents):
        raise RuntimeError("Reranker response does not contain one result per document.")
    try:
        ranked = sorted(results, key=lambda item: float(item["relevance_score"]), reverse=True)
        indices = [int(item["index"]) for item in ranked]
    except (KeyError, TypeError, ValueError) as error:
        raise RuntimeError("Reranker response has an invalid results format.") from error
    if set(indices) != set(range(len(documents))):
        raise RuntimeError("Reranker response contains missing or invalid document indexes.")
    return indices


def calculate_metrics(ranked_lists: list[list[int]], relevant_indices: list[int]) -> dict[str, float]:
    recall_at_1 = recall_at_5 = 0
    reciprocal_rank_at_10 = 0.0
    for ranked_indices, relevant_index in zip(ranked_lists, relevant_indices, strict=True):
        rank = ranked_indices.index(relevant_index) + 1
        recall_at_1 += rank <= 1
        recall_at_5 += rank <= 5
        if rank <= 10:
            reciprocal_rank_at_10 += 1.0 / rank
    query_count = len(relevant_indices)
    return {
        "recall@1": recall_at_1 / query_count,
        "recall@5": recall_at_5 / query_count,
        "mrr@10": reciprocal_rank_at_10 / query_count,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Evaluate a vLLM reranker against test positive documents only."
    )
    parser.add_argument("test_file", type=Path)
    parser.add_argument("model", nargs="?", default="BAAI/bge-reranker-v2-m3")
    parser.add_argument("--url", default=DEFAULT_URL)
    parser.add_argument("--output-dir", type=Path, default=None)
    parser.add_argument("--timeout", type=int, default=120)
    args = parser.parse_args()
    if not args.test_file.is_file():
        parser.error(f"Test file not found: {args.test_file}")
    if args.timeout < 1:
        parser.error("--timeout must be at least 1.")

    started_at = datetime.now().astimezone()
    started_timer = time.perf_counter()
    records = load_test_records(args.test_file)
    corpus, relevant_indices = build_corpus(records, "positive_only")
    ranked_lists = [
        rerank(args.url, args.model, record["query"], corpus, args.timeout)
        for record in tqdm(records, desc="Reranking queries", unit=" queries")
    ]
    output_dir = resolve_output_dir(args.output_dir, "positive_only", args.model, started_at)
    finished_at = datetime.now().astimezone()
    report: dict[str, Any] = {
        "model": args.model,
        "test_file": str(args.test_file.resolve()),
        "corpus": "positive_only",
        "query_count": len(records),
        "document_count": len(corpus),
        "metrics": calculate_metrics(ranked_lists, relevant_indices),
        "generated_at": finished_at.isoformat(timespec="seconds"),
        "started_at": started_at.isoformat(timespec="seconds"),
        "runtime_seconds": time.perf_counter() - started_timer,
        "settings": {"reranker_url": args.url, "timeout_seconds": args.timeout},
        "environment": {"python": sys.version.split()[0], "platform": platform.platform()},
    }
    save_reports(report, output_dir)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
