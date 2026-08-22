"""Evaluate dense retrieval against the unique positive documents in the test set."""

from __future__ import annotations

import argparse
import html
import importlib.metadata
import json
import platform
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

import numpy as np
from tqdm import tqdm


CorpusMode = Literal["positive_only", "positive_and_negative"]


def load_test_records(test_path: Path) -> list[dict[str, Any]]:
    """Load and validate BGE-formatted test records."""
    records: list[dict[str, Any]] = []
    seen_queries: set[str] = set()

    with test_path.open("r", encoding="utf-8") as source:
        for line_number, line in enumerate(
            tqdm(source, desc="Loading test data", unit=" records"), start=1
        ):
            if not line.strip():
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError as error:
                raise ValueError(f"Invalid JSON at line {line_number}: {error}") from error

            query = record.get("query")
            positives = record.get("pos")
            negatives = record.get("neg")
            if not isinstance(query, str) or not query.strip():
                raise ValueError(f"Invalid query at line {line_number}.")
            if (
                not isinstance(positives, list)
                or len(positives) != 1
                or not isinstance(positives[0], str)
                or not positives[0].strip()
            ):
                raise ValueError(f"Exactly one positive is required at line {line_number}.")
            if not isinstance(negatives, list) or not all(
                isinstance(text, str) and text.strip() for text in negatives
            ):
                raise ValueError(f"Invalid negatives at line {line_number}.")

            query = query.strip()
            if query in seen_queries:
                raise ValueError(f"Duplicate query at line {line_number}.")
            seen_queries.add(query)
            records.append(
                {
                    "query": query,
                    "pos": [positives[0].strip()],
                    "neg": [text.strip() for text in negatives],
                }
            )

    if not records:
        raise ValueError("The test file contains no records.")
    return records


def build_corpus(
    records: list[dict[str, Any]], mode: CorpusMode
) -> tuple[list[str], list[int]]:
    """Build a deduplicated corpus and the relevant document index per query."""
    documents: list[str] = []
    document_to_index: dict[str, int] = {}

    def add_document(text: str) -> int:
        if text not in document_to_index:
            document_to_index[text] = len(documents)
            documents.append(text)
        return document_to_index[text]

    relevant_indices = [
        add_document(record["pos"][0])
        for record in tqdm(records, desc="Adding positive documents", unit=" records")
    ]

    if mode == "positive_and_negative":
        for record in tqdm(
            records, desc="Adding negative documents", unit=" records"
        ):
            for negative in record["neg"]:
                add_document(negative)

    return documents, relevant_indices


def normalize_rows(embeddings: np.ndarray) -> np.ndarray:
    """L2-normalize embeddings for inner-product retrieval."""
    embeddings = np.asarray(embeddings, dtype=np.float32)
    norms = np.linalg.norm(embeddings, axis=1, keepdims=True)
    if np.any(norms == 0):
        raise ValueError("The model produced a zero-length embedding.")
    return embeddings / norms


def calculate_metrics(
    query_embeddings: np.ndarray,
    corpus_embeddings: np.ndarray,
    relevant_indices: list[int],
    search_batch_size: int,
) -> dict[str, float]:
    """Calculate Recall@1, Recall@5, and MRR@10 with batched exact search."""
    if len(query_embeddings) != len(relevant_indices):
        raise ValueError("The number of query embeddings and qrels must match.")

    top_k = min(10, len(corpus_embeddings))
    recall_at_1 = 0
    recall_at_5 = 0
    reciprocal_rank_at_10 = 0.0
    corpus_transposed = np.ascontiguousarray(corpus_embeddings.T)

    starts = range(0, len(query_embeddings), search_batch_size)
    for start in tqdm(starts, desc="Exact dense search", unit=" batches"):
        end = min(start + search_batch_size, len(query_embeddings))
        scores = query_embeddings[start:end] @ corpus_transposed

        if top_k == len(corpus_embeddings):
            candidate_indices = np.argsort(-scores, axis=1)[:, :top_k]
        else:
            candidate_indices = np.argpartition(
                -scores, kth=top_k - 1, axis=1
            )[:, :top_k]
            candidate_scores = np.take_along_axis(scores, candidate_indices, axis=1)
            order = np.argsort(-candidate_scores, axis=1)
            candidate_indices = np.take_along_axis(candidate_indices, order, axis=1)

        for offset, ranked_indices in enumerate(candidate_indices):
            relevant_index = relevant_indices[start + offset]
            matches = np.flatnonzero(ranked_indices == relevant_index)
            if not len(matches):
                continue
            rank = int(matches[0]) + 1
            recall_at_1 += rank <= 1
            recall_at_5 += rank <= 5
            reciprocal_rank_at_10 += 1.0 / rank

    query_count = len(query_embeddings)
    return {
        "recall@1": recall_at_1 / query_count,
        "recall@5": recall_at_5 / query_count,
        "mrr@10": reciprocal_rank_at_10 / query_count,
    }


def safe_name(value: str) -> str:
    """Convert a model name or path to a filesystem-safe short name."""
    short_name = value.rstrip("/\\").replace("\\", "/").split("/")[-1]
    return re.sub(r"[^A-Za-z0-9._-]+", "-", short_name).strip("-_") or "model"


def resolve_output_dir(
    requested_dir: Path | None,
    mode: CorpusMode,
    model_name_or_path: str,
    started_at: datetime,
) -> Path:
    """Return the requested directory or a timestamped default directory."""
    if requested_dir is not None:
        return requested_dir
    timestamp = started_at.strftime("%Y%m%d_%H%M%S")
    return (
        Path("reports/evaluations")
        / f"{timestamp}_{mode}_{safe_name(model_name_or_path)}"
    )


def metric_label(metric_name: str) -> str:
    return {"recall@1": "Recall@1", "recall@5": "Recall@5", "mrr@10": "MRR@10"}[
        metric_name
    ]


def build_text_report(report: dict[str, Any]) -> str:
    """Render a compact plain-text evaluation report."""
    metrics = report["metrics"]
    lines = [
        "BGE-M3 Dense Retrieval Evaluation",
        "=" * 33,
        f"Generated at: {report['generated_at']}",
        f"Model: {report['model']}",
        f"Corpus: {report['corpus']}",
        f"Test file: {report['test_file']}",
        f"Queries: {report['query_count']:,}",
        f"Documents: {report['document_count']:,}",
        f"Runtime: {report['runtime_seconds']:.2f} seconds",
        "",
        "Metrics",
        "-------",
    ]
    lines.extend(
        f"{metric_label(name)}: {value:.6f} ({value:.2%})"
        for name, value in metrics.items()
    )
    lines.extend(
        [
            "",
            "Settings",
            "--------",
            f"Devices: {', '.join(report['settings']['devices'])}",
            f"FP16: {report['settings']['fp16']}",
            f"Encode batch size: {report['settings']['encode_batch_size']}",
            f"Search batch size: {report['settings']['search_batch_size']}",
            f"Query max length: {report['settings']['query_max_length']}",
            f"Passage max length: {report['settings']['passage_max_length']}",
            "",
            "Environment",
            "-----------",
            f"Python: {report['environment']['python']}",
            f"Platform: {report['environment']['platform']}",
            f"FlagEmbedding: {report['environment']['flag_embedding']}",
        ]
    )
    return "\n".join(lines) + "\n"


def build_html_report(report: dict[str, Any]) -> str:
    """Render a standalone professional HTML evaluation report."""
    escape = lambda value: html.escape(str(value))
    metric_cards = "".join(
        f"""
        <article class="metric-card">
          <span>{escape(metric_label(name))}</span>
          <strong>{value:.2%}</strong>
          <small>{value:.6f}</small>
        </article>"""
        for name, value in report["metrics"].items()
    )
    settings_rows = "".join(
        f"<tr><th>{escape(key.replace('_', ' ').title())}</th><td>{escape(value)}</td></tr>"
        for key, value in {
            **report["settings"],
            **report["environment"],
        }.items()
    )
    corpus_title = (
        "Positive-only corpus"
        if report["corpus"] == "positive_only"
        else "Positive + negative corpus"
    )
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>BGE-M3 Evaluation Report</title>
  <style>
    :root {{ color-scheme: dark; --bg:#07111f; --panel:#0e1b2d; --line:#203653;
      --text:#e9f1fb; --muted:#91a4bc; --accent:#53d6b5; --blue:#65a8ff; }}
    * {{ box-sizing:border-box; }}
    body {{ margin:0; background:radial-gradient(circle at top right,#16345a 0,#07111f 42%);
      color:var(--text); font:15px/1.6 Inter,Segoe UI,Arial,sans-serif; min-height:100vh; }}
    main {{ width:min(1080px,calc(100% - 32px)); margin:42px auto; }}
    header {{ padding:34px; border:1px solid var(--line); border-radius:22px;
      background:linear-gradient(135deg,rgba(14,27,45,.96),rgba(10,22,39,.88));
      box-shadow:0 24px 70px rgba(0,0,0,.3); }}
    .eyebrow {{ color:var(--accent); text-transform:uppercase; letter-spacing:.14em;
      font-size:12px; font-weight:700; }}
    h1 {{ margin:8px 0 5px; font-size:clamp(28px,5vw,48px); line-height:1.1; }}
    .subtitle {{ color:var(--muted); margin:0; }}
    .badge {{ display:inline-block; margin-top:20px; padding:7px 12px; border-radius:999px;
      background:rgba(83,214,181,.12); color:var(--accent); border:1px solid rgba(83,214,181,.35); }}
    .metrics {{ display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin:22px 0; }}
    .metric-card,.panel {{ background:rgba(14,27,45,.92); border:1px solid var(--line);
      border-radius:18px; box-shadow:0 14px 35px rgba(0,0,0,.2); }}
    .metric-card {{ padding:24px; }} .metric-card span,.metric-card small {{ color:var(--muted); }}
    .metric-card strong {{ display:block; color:var(--accent); font-size:36px; margin:5px 0 0; }}
    .panel {{ padding:26px; margin-top:16px; }} h2 {{ margin:0 0 17px; font-size:20px; }}
    .facts {{ display:grid; grid-template-columns:repeat(2,1fr); gap:12px; }}
    .fact {{ padding:14px 16px; border-radius:12px; background:#0a1728; }}
    .fact span {{ display:block; color:var(--muted); font-size:12px; text-transform:uppercase; }}
    .fact strong {{ display:block; overflow-wrap:anywhere; }}
    table {{ width:100%; border-collapse:collapse; }} th,td {{ padding:11px 8px; border-bottom:1px solid var(--line); text-align:left; }}
    th {{ width:34%; color:var(--muted); font-weight:500; }} footer {{ color:var(--muted); text-align:center; margin:24px; }}
    @media(max-width:700px) {{ .metrics,.facts {{ grid-template-columns:1fr; }} main {{ margin:20px auto; }} }}
  </style>
</head>
<body>
<main>
  <header>
    <div class="eyebrow">Dense Retrieval Benchmark</div>
    <h1>BGE-M3 Evaluation</h1>
    <p class="subtitle">A reproducible retrieval baseline generated from the held-out test split.</p>
    <div class="badge">{escape(corpus_title)}</div>
  </header>
  <section class="metrics">{metric_cards}</section>
  <section class="panel">
    <h2>Evaluation summary</h2>
    <div class="facts">
      <div class="fact"><span>Model</span><strong>{escape(report['model'])}</strong></div>
      <div class="fact"><span>Generated</span><strong>{escape(report['generated_at'])}</strong></div>
      <div class="fact"><span>Queries</span><strong>{report['query_count']:,}</strong></div>
      <div class="fact"><span>Documents</span><strong>{report['document_count']:,}</strong></div>
      <div class="fact"><span>Runtime</span><strong>{report['runtime_seconds']:.2f} seconds</strong></div>
      <div class="fact"><span>Test file</span><strong>{escape(report['test_file'])}</strong></div>
    </div>
  </section>
  <section class="panel"><h2>Configuration and environment</h2><table>{settings_rows}</table></section>
  <footer>Generated by the finetune-embedding-models evaluation pipeline</footer>
</main>
</body>
</html>
"""


def save_reports(report: dict[str, Any], output_dir: Path) -> dict[str, str]:
    """Save JSON, plain-text, and standalone HTML reports."""
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "json": output_dir / "report.json",
        "text": output_dir / "report.txt",
        "html": output_dir / "report.html",
    }
    serialized_paths = {name: str(path) for name, path in paths.items()}
    report["report_files"] = serialized_paths
    paths["json"].write_text(
        json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    paths["text"].write_text(build_text_report(report), encoding="utf-8")
    paths["html"].write_text(build_html_report(report), encoding="utf-8")
    return serialized_paths


def evaluate(
    test_path: Path,
    model_name_or_path: str,
    output_dir: Path | None,
    mode: CorpusMode,
    devices: list[str] | None,
    use_fp16: bool,
    encode_batch_size: int,
    search_batch_size: int,
    query_max_length: int,
    passage_max_length: int,
) -> dict[str, Any]:
    """Run one dense-retrieval evaluation and save all report formats."""
    started_at = datetime.now().astimezone()
    started_timer = time.perf_counter()
    resolved_output_dir = resolve_output_dir(
        output_dir, mode, model_name_or_path, started_at
    )
    records = load_test_records(test_path)
    queries = [record["query"] for record in records]
    corpus, relevant_indices = build_corpus(records, mode)

    try:
        from FlagEmbedding import BGEM3FlagModel
    except ImportError as error:
        raise RuntimeError(
            "FlagEmbedding is not installed. Run: pip install -r requirements.txt"
        ) from error

    model = BGEM3FlagModel(
        model_name_or_path,
        use_fp16=use_fp16,
        pooling_method="cls",
        devices=devices,
    )
    query_embeddings = model.encode_queries(
        queries,
        batch_size=encode_batch_size,
        max_length=query_max_length,
        return_dense=True,
        return_sparse=False,
        return_colbert_vecs=False,
    )["dense_vecs"]
    corpus_embeddings = model.encode_corpus(
        corpus,
        batch_size=encode_batch_size,
        max_length=passage_max_length,
        return_dense=True,
        return_sparse=False,
        return_colbert_vecs=False,
    )["dense_vecs"]

    metrics = calculate_metrics(
        normalize_rows(query_embeddings),
        normalize_rows(corpus_embeddings),
        relevant_indices,
        search_batch_size,
    )
    try:
        flag_embedding_version = importlib.metadata.version("FlagEmbedding")
    except importlib.metadata.PackageNotFoundError:
        flag_embedding_version = "unknown"
    finished_at = datetime.now().astimezone()
    report: dict[str, Any] = {
        "model": model_name_or_path,
        "test_file": str(test_path.resolve()),
        "corpus": mode,
        "query_count": len(queries),
        "document_count": len(corpus),
        "metrics": metrics,
        "generated_at": finished_at.isoformat(timespec="seconds"),
        "started_at": started_at.isoformat(timespec="seconds"),
        "runtime_seconds": time.perf_counter() - started_timer,
        "settings": {
            "devices": devices or ["auto"],
            "fp16": use_fp16,
            "encode_batch_size": encode_batch_size,
            "search_batch_size": search_batch_size,
            "query_max_length": query_max_length,
            "passage_max_length": passage_max_length,
        },
        "environment": {
            "python": sys.version.split()[0],
            "platform": platform.platform(),
            "flag_embedding": flag_embedding_version,
        },
    }
    save_reports(report, resolved_output_dir)
    return report


def create_parser(description: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("test_file", type=Path)
    parser.add_argument("model_name_or_path")
    parser.add_argument("--output-dir", type=Path, default=None)
    parser.add_argument("--devices", nargs="+", default=None)
    parser.add_argument("--fp16", action="store_true")
    parser.add_argument("--encode-batch-size", type=int, default=16)
    parser.add_argument("--search-batch-size", type=int, default=64)
    parser.add_argument("--query-max-length", type=int, default=512)
    parser.add_argument("--passage-max-length", type=int, default=512)
    return parser


def validate_args(parser: argparse.ArgumentParser, args: argparse.Namespace) -> None:
    if not args.test_file.is_file():
        parser.error(f"Test file not found: {args.test_file}")
    if args.output_dir is not None and args.test_file.resolve() == args.output_dir.resolve():
        parser.error("Output directory must be different from the test file.")
    for name in (
        "encode_batch_size",
        "search_batch_size",
        "query_max_length",
        "passage_max_length",
    ):
        if getattr(args, name) < 1:
            parser.error(f"--{name.replace('_', '-')} must be at least 1.")


def main() -> None:
    parser = create_parser("Evaluate BGE-M3 against test positive documents only.")
    args = parser.parse_args()
    validate_args(parser, args)
    report = evaluate(
        args.test_file,
        args.model_name_or_path,
        args.output_dir,
        mode="positive_only",
        devices=args.devices,
        use_fp16=args.fp16,
        encode_batch_size=args.encode_batch_size,
        search_batch_size=args.search_batch_size,
        query_max_length=args.query_max_length,
        passage_max_length=args.passage_max_length,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
