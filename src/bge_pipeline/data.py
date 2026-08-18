from __future__ import annotations

import csv
import hashlib
import json
import random
import re
import unicodedata
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

EXPECTED_COLUMNS = ["question", "content_text"]
_SPACE_RE = re.compile(r"\s+")
_PERSIAN_TRANSLATION = str.maketrans({"ي": "ی", "ى": "ی", "ك": "ک"})


def normalize_persian(text: str) -> str:
    """Conservatively normalize Unicode and spacing without removing punctuation."""
    text = unicodedata.normalize("NFC", text).translate(_PERSIAN_TRANSLATION)
    return _SPACE_RE.sub(" ", text).strip()


def stable_id(*parts: str, prefix: str = "") -> str:
    digest = hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()[:20]
    return f"{prefix}{digest}"


def file_sha256(path: Path, chunk_size: int = 1024 * 1024) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(chunk_size), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _percentiles(values: list[int]) -> dict[str, int]:
    values = sorted(values)
    if not values:
        return {}
    result = {}
    for name, fraction in (("min", 0), ("p50", .5), ("p90", .9), ("p95", .95), ("p99", .99), ("max", 1)):
        result[name] = values[round((len(values) - 1) * fraction)]
    return result


def _assign_groups(group_sizes: dict[str, int], seed: int) -> dict[str, str]:
    """Assign whole normalized-question groups near an 80/10/10 row ratio."""
    items = list(group_sizes.items())
    random.Random(seed).shuffle(items)
    items.sort(key=lambda item: item[1], reverse=True)
    total = sum(group_sizes.values())
    targets = {"train": total * .8, "validation": total * .1, "test": total * .1}
    counts = Counter()
    assignments = {}
    for group, size in items:
        split = min(targets, key=lambda name: counts[name] / targets[name])
        assignments[group] = split
        counts[split] += size
    return assignments


def read_and_split(input_path: Path, seed: int = 42) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen_pairs: set[tuple[str, str]] = set()
    raw_rows = empty_questions = empty_answers = duplicate_pairs = 0
    question_counts: Counter[str] = Counter()
    raw_question_counts: Counter[str] = Counter()
    raw_pair_counts: Counter[tuple[str, str]] = Counter()

    with input_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames != EXPECTED_COLUMNS:
            raise ValueError(f"Expected CSV columns {EXPECTED_COLUMNS}, got {reader.fieldnames}")
        for line_number, row in enumerate(reader, start=2):
            raw_rows += 1
            question = (row.get("question") or "").strip()
            answer = (row.get("content_text") or "").strip()
            if not question:
                empty_questions += 1
            if not answer:
                empty_answers += 1
            if not question or not answer:
                raise ValueError(f"Empty question/answer at CSV line {line_number}")
            normalized_question = normalize_persian(question)
            normalized_answer = normalize_persian(answer)
            raw_question_counts[question] += 1
            raw_pair_counts[(question, answer)] += 1
            pair_key = (normalized_question, normalized_answer)
            question_counts[normalized_question] += 1
            if pair_key in seen_pairs:
                duplicate_pairs += 1
                continue
            seen_pairs.add(pair_key)
            rows.append({
                "id": stable_id(normalized_question, normalized_answer, prefix="pair_"),
                "question_id": stable_id(normalized_question, prefix="q_"),
                "document_id": stable_id(normalized_answer, prefix="doc_"),
                "question": question,
                "content_text": answer,
                "normalized_question": normalized_question,
            })

    groups = Counter(row["normalized_question"] for row in rows)
    assignments = _assign_groups(dict(groups), seed)
    for row in rows:
        row["split"] = assignments[row["normalized_question"]]

    split_counts = Counter(row["split"] for row in rows)
    report = {
        "input": str(input_path),
        "input_sha256": file_sha256(input_path),
        "seed": seed,
        "raw_rows": raw_rows,
        "written_rows": len(rows),
        "empty_questions": empty_questions,
        "empty_answers": empty_answers,
        "duplicate_pair_rows_removed": duplicate_pairs,
        "duplicate_question_extra_rows_raw": sum(max(0, count - 1) for count in raw_question_counts.values()),
        "duplicate_pair_extra_rows_raw": sum(max(0, count - 1) for count in raw_pair_counts.values()),
        "duplicate_question_extra_rows_after_normalization": sum(max(0, count - 1) for count in question_counts.values()),
        "unique_normalized_questions": len(groups),
        "split_counts": dict(split_counts),
        "question_characters": _percentiles([len(row["question"]) for row in rows]),
        "answer_characters": _percentiles([len(row["content_text"]) for row in rows]),
    }
    return rows, report


def add_token_statistics(rows: list[dict[str, Any]], report: dict[str, Any], model_name: str) -> None:
    from transformers import AutoTokenizer

    tokenizer = AutoTokenizer.from_pretrained(model_name)
    question_lengths: list[int] = []
    answer_lengths: list[int] = []
    for row in rows:
        question_lengths.append(len(tokenizer(row["question"], add_special_tokens=True, truncation=False)["input_ids"]))
        answer_lengths.append(len(tokenizer(row["content_text"], add_special_tokens=True, truncation=False)["input_ids"]))
    report["tokenizer"] = model_name
    report["question_tokens"] = _percentiles(question_lengths)
    report["answer_tokens"] = _percentiles(answer_lengths)
    report["answers_over_8192_tokens"] = sum(length > 8192 for length in answer_lengths)


def write_outputs(rows: list[dict[str, Any]], report: dict[str, Any], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    public_columns = ["id", "question_id", "document_id", "question", "content_text", "split"]
    for split in ("train", "validation", "test"):
        split_rows = [{key: row[key] for key in public_columns} for row in rows if row["split"] == split]
        jsonl_path = output_dir / f"{split}.jsonl"
        with jsonl_path.open("w", encoding="utf-8", newline="\n") as handle:
            for row in split_rows:
                handle.write(json.dumps(row, ensure_ascii=False) + "\n")
        try:
            import pyarrow as pa
            import pyarrow.parquet as pq
            pq.write_table(pa.Table.from_pylist(split_rows), output_dir / f"{split}.parquet", compression="zstd")
        except ImportError as exc:
            raise RuntimeError("pyarrow is required to produce the requested Parquet outputs") from exc
    (output_dir / "manifest.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")


def load_jsonl(path: Path) -> list[dict[str, Any]]:
    with path.open("r", encoding="utf-8") as handle:
        return [json.loads(line) for line in handle if line.strip()]


def export_duplicate_question_extras(input_path: Path, output_path: Path) -> dict[str, int]:
    """Export question occurrences after the first; flag exact pair repetitions."""
    seen_questions: Counter[str] = Counter()
    seen_pairs: Counter[tuple[str, str]] = Counter()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    exported = exact_pair_extras = 0
    fieldnames = [
        "source_line", "question", "content_text", "normalized_question",
        "question_occurrence", "pair_occurrence", "is_exact_pair_duplicate",
    ]
    with input_path.open("r", encoding="utf-8-sig", newline="") as source, output_path.open(
        "w", encoding="utf-8-sig", newline=""
    ) as destination:
        reader = csv.DictReader(source)
        if reader.fieldnames != EXPECTED_COLUMNS:
            raise ValueError(f"Expected CSV columns {EXPECTED_COLUMNS}, got {reader.fieldnames}")
        writer = csv.DictWriter(destination, fieldnames=fieldnames)
        writer.writeheader()
        for line_number, row in enumerate(reader, start=2):
            question = (row.get("question") or "").strip()
            answer = (row.get("content_text") or "").strip()
            normalized_question = normalize_persian(question)
            # This audit intentionally matches the raw-data counts (1,624/258):
            # whitespace is trimmed, but Persian glyph variants are not collapsed.
            question_key = question
            pair = (question, answer)
            seen_questions[question_key] += 1
            seen_pairs[pair] += 1
            if seen_questions[question_key] == 1:
                continue
            is_exact = seen_pairs[pair] > 1
            exact_pair_extras += int(is_exact)
            exported += 1
            writer.writerow({
                "source_line": line_number,
                "question": question,
                "content_text": answer,
                "normalized_question": normalized_question,
                "question_occurrence": seen_questions[question_key],
                "pair_occurrence": seen_pairs[pair],
                "is_exact_pair_duplicate": is_exact,
            })
    return {"duplicate_question_extra_rows": exported, "exact_pair_duplicate_extra_rows": exact_pair_extras}


def assert_split_integrity(output_dir: Path) -> None:
    seen_questions: dict[str, str] = {}
    seen_pairs: set[str] = set()
    for split in ("train", "validation", "test"):
        for row in load_jsonl(output_dir / f"{split}.jsonl"):
            normalized = normalize_persian(row["question"])
            previous = seen_questions.setdefault(normalized, split)
            if previous != split:
                raise AssertionError(f"Question leakage between {previous} and {split}: {row['question_id']}")
            if row["id"] in seen_pairs:
                raise AssertionError(f"Duplicate pair id: {row['id']}")
            seen_pairs.add(row["id"])
