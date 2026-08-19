from __future__ import annotations

import argparse
import csv
import json
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Iterator

from openai import OpenAI
from tqdm import tqdm


SYSTEM_PROMPT = """/no_think
You are an expert in information retrieval and contrastive training for text embedding models.
Generate exactly 7 high-quality hard negatives for the supplied query-positive pair.

First infer the exact information need, why the positive satisfies it, and the critical requirements. A hard negative must be strongly related, lexically and semantically similar, natural and corpus-like, and approximately the same style and length as the positive, but it must NOT satisfy the actual information need or contain the correct answer in another form. Whenever possible, violate only one critical requirement.

Use exactly these seven strategies in this order:
1 entity_confusion: replace one critical entity with a plausible related entity.
2 attribute_confusion: alter a critical property, number, date, value, or attribute.
3 relation_confusion: preserve entities and terminology but change their relationship.
4 intent_drift: answer a neighboring question about almost the same subject.
5 constraint_violation: violate one time, location, scope, condition, version, population, or category constraint.
6 partial_relevance: provide useful context but omit the specifically required information.
7 lexical_trap: use highly similar wording and structure with a subtly different meaning.

Internally reject any candidate that could answer the query, is trivially unrelated, is insufficiently similar, differs only in wording, or sounds artificial.

Special rule for list/set/broad queries: if the query accepts multiple valid entities, sources, examples, methods, dates, or items, replacing the positive item with another valid member of that set is a FALSE NEGATIVE. Never do that. The changed entity or value must make the passage fail an explicit requirement of the query. Likewise, a passage that supplies any independently valid requested source/example is still positive and must be rejected. Before emitting each item, explicitly test it against the original query—not merely against the positive document—and rewrite it if a user could reasonably accept it as an answer.

Return ONLY one valid JSON object (no markdown or commentary) with this exact shape:
{
  "query": "...",
  "information_need": "...",
  "critical_requirements": ["..."],
  "hard_negatives": [
    {"id": 1, "type": "entity_confusion", "hardness": "very_hard", "violated_requirement": "...", "text": "..."},
    {"id": 2, "type": "attribute_confusion", "hardness": "very_hard", "violated_requirement": "...", "text": "..."},
    {"id": 3, "type": "relation_confusion", "hardness": "very_hard", "violated_requirement": "...", "text": "..."},
    {"id": 4, "type": "intent_drift", "hardness": "hard", "violated_requirement": "...", "text": "..."},
    {"id": 5, "type": "constraint_violation", "hardness": "hard", "violated_requirement": "...", "text": "..."},
    {"id": 6, "type": "partial_relevance", "hardness": "medium_hard", "violated_requirement": "...", "text": "..."},
    {"id": 7, "type": "lexical_trap", "hardness": "hard", "violated_requirement": "...", "text": "..."}
  ]
}"""

EXPECTED_TYPES = [
    "entity_confusion", "attribute_confusion", "relation_confusion",
    "intent_drift", "constraint_violation", "partial_relevance", "lexical_trap",
]
_local = threading.local()


def client(base_url: str, api_key: str, timeout: float) -> OpenAI:
    key = (base_url, api_key, timeout)
    if getattr(_local, "key", None) != key:
        _local.client = OpenAI(base_url=base_url, api_key=api_key, timeout=timeout)
        _local.key = key
    return _local.client


def parse_json(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
    try:
        value = json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start < 0 or end <= start:
            raise
        value = json.loads(text[start:end + 1])
    if not isinstance(value, dict):
        raise ValueError("response is not a JSON object")
    return value


def validate(value: dict[str, Any]) -> None:
    if not isinstance(value.get("information_need"), str) or not value["information_need"].strip():
        raise ValueError("missing information_need")
    requirements = value.get("critical_requirements")
    if not isinstance(requirements, list) or not requirements:
        raise ValueError("critical_requirements must be a non-empty list")
    negatives = value.get("hard_negatives")
    if not isinstance(negatives, list) or len(negatives) != 7:
        raise ValueError("hard_negatives must contain exactly 7 items")
    for i, (item, expected_type) in enumerate(zip(negatives, EXPECTED_TYPES), 1):
        if not isinstance(item, dict) or item.get("id") != i or item.get("type") != expected_type:
            raise ValueError(f"invalid hard negative #{i}")
        if not isinstance(item.get("text"), str) or not item["text"].strip():
            raise ValueError(f"empty text in hard negative #{i}")
        if not isinstance(item.get("violated_requirement"), str) or not item["violated_requirement"].strip():
            raise ValueError(f"missing violated_requirement in hard negative #{i}")


def generate(row: tuple[int, str, str], args: argparse.Namespace) -> dict[str, Any]:
    source_index, query, positive = row
    user_prompt = f"Query:\n{query}\n\nPositive document:\n{positive}"
    last_error: Exception | None = None
    for attempt in range(1, args.retries + 1):
        try:
            response = client(args.base_url, args.api_key, args.timeout).chat.completions.create(
                model=args.model,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
                temperature=args.temperature,
                max_tokens=args.max_tokens,
                extra_body={
                    "chat_template_kwargs": {"enable_thinking": False},
                },
            )
            content = response.choices[0].message.content
            if not content:
                raise ValueError("model returned empty content")
            value = parse_json(content)
            validate(value)
            # The source text is authoritative; do not trust the model to copy it exactly.
            value["query"] = query
            return {
                "source_index": source_index,
                "query": query,
                "positive_document": positive,
                "information_need": value["information_need"],
                "critical_requirements": value["critical_requirements"],
                "hard_negatives": value["hard_negatives"],
            }
        except Exception as exc:  # network, server, malformed JSON, or failed validation
            last_error = exc
            if attempt < args.retries:
                time.sleep(args.retry_delay * (2 ** (attempt - 1)) + random.random())
    raise RuntimeError(f"row {source_index} failed after {args.retries} attempts: {last_error}")


def repair_output_tail(path: Path) -> None:
    """Remove only a crash-truncated final line, or add its missing newline."""
    if not path.exists() or path.stat().st_size == 0:
        return
    with path.open("rb+") as handle:
        handle.seek(0, os.SEEK_END)
        end = handle.tell()
        handle.seek(end - 1)
        if handle.read(1) == b"\n":
            return
        position = end - 1
        while position > 0:
            position -= 1
            handle.seek(position)
            if handle.read(1) == b"\n":
                position += 1
                break
        handle.seek(position)
        tail = handle.read(end - position)
        try:
            json.loads(tail.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            handle.truncate(position)
            print("Warning: removed a crash-truncated final output line")
        else:
            handle.seek(0, os.SEEK_END)
            handle.write(b"\n")
        handle.flush()
        os.fsync(handle.fileno())


def completed_indices(path: Path) -> set[int]:
    done: set[int] = set()
    if not path.exists():
        return done
    with path.open("r", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            try:
                value = json.loads(line)
                done.add(int(value["source_index"]))
            except (json.JSONDecodeError, KeyError, TypeError, ValueError):
                print(f"Warning: ignoring incomplete/invalid output line {line_number}")
    return done


def input_rows(path: Path, done: set[int], limit: int | None) -> Iterator[tuple[int, str, str]]:
    yielded = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"question", "content_text"}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError(f"CSV must contain columns {sorted(required)}")
        for source_index, item in enumerate(reader):
            if source_index in done:
                continue
            query = (item["question"] or "").strip()
            positive = (item["content_text"] or "").strip()
            if not query or not positive:
                continue
            yield source_index, query, positive
            yielded += 1
            if limit is not None and yielded >= limit:
                return


def chunks(rows: Iterator[tuple[int, str, str]], size: int) -> Iterator[list[tuple[int, str, str]]]:
    while True:
        batch = []
        try:
            for _ in range(size):
                batch.append(next(rows))
        except StopIteration:
            pass
        if not batch:
            return
        yield batch


def append_line(handle: Any, value: dict[str, Any]) -> None:
    handle.write(json.dumps(value, ensure_ascii=False) + "\n")
    handle.flush()
    os.fsync(handle.fileno())


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate resumable hard negatives with parallel OpenAI-compatible requests.")
    parser.add_argument("--input", type=Path, default=Path("porseman_clean.csv"))
    parser.add_argument("--output", type=Path, default=Path("data/hard_negatives.jsonl"))
    parser.add_argument("--base-url", default="http://127.0.0.1:18000/v1")
    parser.add_argument("--api-key", default="dummy")
    parser.add_argument("--model", default="Qwen/Qwen3.5-397B-A17B-FP8")
    parser.add_argument("--batch-size", type=int, default=50)
    parser.add_argument("--workers", type=int, default=50)
    parser.add_argument("--retries", type=int, default=4)
    parser.add_argument("--retry-delay", type=float, default=2.0)
    parser.add_argument("--timeout", type=float, default=300.0)
    parser.add_argument("--temperature", type=float, default=0.3)
    parser.add_argument("--max-tokens", type=int, default=8192)
    parser.add_argument("--limit", type=int, help="Process at most this many unfinished rows (useful for testing).")
    args = parser.parse_args()
    if min(args.batch_size, args.workers, args.retries) < 1:
        parser.error("batch-size, workers, and retries must be positive")

    args.output.parent.mkdir(parents=True, exist_ok=True)
    repair_output_tail(args.output)
    error_path = args.output.with_suffix(args.output.suffix + ".errors.jsonl")
    done = completed_indices(args.output)
    rows = input_rows(args.input, done, args.limit)
    pending = sum(1 for _ in input_rows(args.input, done, args.limit))
    succeeded = failed = 0
    print(f"Resuming with {len(done)} completed rows; pending={pending}; output={args.output}")

    try:
        with args.output.open("a", encoding="utf-8", newline="\n") as output, \
                error_path.open("a", encoding="utf-8", newline="\n") as errors, \
                ThreadPoolExecutor(max_workers=args.workers) as executor, \
                tqdm(total=pending, unit="row", dynamic_ncols=True, desc="Hard negatives") as progress:
            for batch_number, batch in enumerate(chunks(rows, args.batch_size), 1):
                futures = {executor.submit(generate, row, args): row[0] for row in batch}
                for future in as_completed(futures):
                    source_index = futures[future]
                    try:
                        append_line(output, future.result())
                        succeeded += 1
                    except Exception as exc:
                        append_line(errors, {"source_index": source_index, "error": str(exc), "time": time.time()})
                        failed += 1
                    progress.update(1)
                    progress.set_postfix(saved=succeeded, failed=failed, refresh=False)
    except KeyboardInterrupt:
        print("Interrupted safely. Run the same command to resume.")
        raise SystemExit(130)

    print(f"Finished: saved={succeeded}, failed={failed}. Failed rows will be retried on the next run.")


if __name__ == "__main__":
    main()
