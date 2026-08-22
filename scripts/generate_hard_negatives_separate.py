"""Generate seven hard negatives with one independent LLM request per strategy.

Each successful negative is committed to a SQLite checkpoint before the next
result is processed. Re-running the same command therefore resumes incomplete
rows and incomplete strategy sets without regenerating saved work.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import random
import re
import sqlite3
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator

from openai import OpenAI
from tqdm import tqdm


PROMPT_VERSION = "separate-v1"
STRATEGIES = (
    {
        "type": "entity_confusion",
        "hardness": "very_hard",
        "instruction": (
            "Replace one critical entity with a plausible, closely related entity. "
            "The replacement must make the passage fail the query, not merely give "
            "another valid member of a requested list or set."
        ),
    },
    {
        "type": "attribute_confusion",
        "hardness": "very_hard",
        "instruction": (
            "Alter one critical property, number, date, value, condition, or attribute "
            "while preserving the surrounding topic and wording."
        ),
    },
    {
        "type": "relation_confusion",
        "hardness": "very_hard",
        "instruction": (
            "Preserve the important entities and terminology but change one critical "
            "relationship, direction, cause, consequence, or role."
        ),
    },
    {
        "type": "intent_drift",
        "hardness": "hard",
        "instruction": (
            "Answer a neighboring question about almost the same subject while not "
            "satisfying the user's actual information need."
        ),
    },
    {
        "type": "constraint_violation",
        "hardness": "hard",
        "instruction": (
            "Violate exactly one important time, location, scope, condition, version, "
            "population, category, or other explicit constraint in the query."
        ),
    },
    {
        "type": "partial_relevance",
        "hardness": "medium_hard",
        "instruction": (
            "Provide strongly relevant and useful context but omit the particular fact "
            "or conclusion required to answer the query."
        ),
    },
    {
        "type": "lexical_trap",
        "hardness": "hard",
        "instruction": (
            "Use wording and structure highly similar to the positive passage while "
            "introducing a subtle semantic difference that makes it an incorrect answer."
        ),
    },
)

BASE_SYSTEM_PROMPT = """You are an expert in information retrieval and contrastive training for text embedding models.

Create exactly ONE high-quality hard negative for the supplied query-positive pair using only the assigned strategy. Independently infer the user's exact information need and the critical requirements before writing, but do not output that analysis.

A hard negative must be strongly related, lexically and semantically similar, natural, corpus-like, and in the same language and style as the positive passage, while NOT satisfying the query and NOT containing the correct answer in another form. Prefer changing or violating only one critical requirement. Do not invent irrelevant text, padding, labels, or commentary.

For broad, list, set, example, source, or multi-answer queries, another valid item is still a positive answer and is forbidden. Test the candidate against the original query, not just against the positive passage. Reject and rewrite it internally if a reasonable user could accept it as an answer.

LENGTH IS A HIGHEST-PRIORITY REQUIREMENT. The target and permitted word-count interval are supplied by the user. Aim for exactly the target word count. If exact equality is not possible, remain strictly inside the permitted interval and as close to the target as possible. Count whitespace-delimited words in the finished passage before responding. Never use filler, repetition, metadata, or incomplete sentences to adjust length.

Return ONLY one valid JSON object, without Markdown or commentary, in this exact shape:
{"violated_requirement":"a concise description","text":"the complete hard-negative passage"}
"""

_local = threading.local()


@dataclass(frozen=True)
class SourceRow:
    source_index: int
    query: str
    positive: str


def system_prompt(enable_thinking: bool) -> str:
    """Return the common instructions for a single-strategy request."""
    return BASE_SYSTEM_PROMPT if enable_thinking else "/no_think\n" + BASE_SYSTEM_PROMPT


def get_client(base_url: str, api_key: str, timeout: float) -> OpenAI:
    """Create one reusable OpenAI client per worker thread."""
    key = (base_url, api_key, timeout)
    if getattr(_local, "key", None) != key:
        _local.client = OpenAI(base_url=base_url, api_key=api_key, timeout=timeout)
        _local.key = key
    return _local.client


def word_bounds(target: int, tolerance: float) -> tuple[int, int]:
    """Return an inclusive, symmetric word-count interval around the target."""
    margin = math.floor(target * tolerance)
    return max(1, target - margin), target + margin


def parse_json(text: str) -> dict[str, Any]:
    """Parse a JSON object, tolerating an accidental Markdown fence."""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I)
    try:
        value = json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start < 0 or end <= start:
            raise
        value = json.loads(text[start : end + 1])
    if not isinstance(value, dict):
        raise ValueError("response is not a JSON object")
    return value


def validate_candidate(
    value: dict[str, Any], positive: str, minimum_words: int, maximum_words: int
) -> dict[str, str]:
    """Validate and normalize one generated hard negative."""
    text = value.get("text")
    violated = value.get("violated_requirement")
    if not isinstance(text, str) or not text.strip():
        raise ValueError("missing or empty text")
    if not isinstance(violated, str) or not violated.strip():
        raise ValueError("missing violated_requirement")

    text = text.strip()
    if text == positive.strip():
        raise ValueError("the generated negative is identical to the positive passage")
    count = len(text.split())
    if not minimum_words <= count <= maximum_words:
        raise ValueError(
            f"generated passage has {count} words; expected {minimum_words}..{maximum_words}"
        )
    return {"violated_requirement": violated.strip(), "text": text}


def make_user_prompt(row: SourceRow, strategy_index: int, tolerance: float) -> str:
    """Build a fully independent prompt for exactly one strategy."""
    strategy = STRATEGIES[strategy_index]
    target = len(row.positive.split())
    minimum, maximum = word_bounds(target, tolerance)
    return (
        f"Assigned strategy: {strategy['type']}\n"
        f"Strategy rule: {strategy['instruction']}\n\n"
        f"Query:\n{row.query}\n\n"
        f"Positive passage:\n{row.positive}\n\n"
        "MANDATORY LENGTH CHECK:\n"
        f"- Target: exactly {target} whitespace-delimited words.\n"
        f"- Acceptable only when exact equality is impossible: {minimum} to {maximum} words inclusive.\n"
        "- Exact equality is strongly preferred. Count again before returning JSON.\n\n"
        "Generate only this one strategy. Return only the required JSON object."
    )


def generate_candidate(
    row: SourceRow, strategy_index: int, args: argparse.Namespace
) -> dict[str, str]:
    """Generate and validate one strategy, retrying only this request on failure."""
    target = len(row.positive.split())
    minimum, maximum = word_bounds(target, args.length_tolerance)
    prompt = make_user_prompt(row, strategy_index, args.length_tolerance)
    last_error: Exception | None = None

    for attempt in range(1, args.retries + 1):
        try:
            response = get_client(args.base_url, args.api_key, args.timeout).chat.completions.create(
                model=args.model,
                messages=[
                    {"role": "system", "content": system_prompt(args.enable_thinking)},
                    {"role": "user", "content": prompt},
                ],
                temperature=args.temperature,
                max_tokens=args.max_tokens,
                extra_body={
                    "chat_template_kwargs": {"enable_thinking": args.enable_thinking},
                },
            )
            content = response.choices[0].message.content
            if not content:
                raise ValueError("model returned empty content")
            return validate_candidate(parse_json(content), row.positive, minimum, maximum)
        except Exception as exc:  # network, malformed JSON, validation, or server error
            last_error = exc
            if attempt < args.retries:
                time.sleep(args.retry_delay * (2 ** (attempt - 1)) + random.random())

    strategy_type = STRATEGIES[strategy_index]["type"]
    raise RuntimeError(
        f"row {row.source_index}, strategy {strategy_type} failed after "
        f"{args.retries} attempts: {last_error}"
    )


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
    """Load source indexes already committed to the final JSONL file."""
    done: set[int] = set()
    if not path.exists():
        return done
    with path.open("r", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle, 1):
            if not line.strip():
                continue
            try:
                value = json.loads(line)
                source_index = int(value["source_index"])
            except (json.JSONDecodeError, KeyError, TypeError, ValueError) as exc:
                raise ValueError(
                    f"invalid completed output record at line {line_number}: {exc}"
                ) from exc
            if source_index in done:
                raise ValueError(f"duplicate source_index {source_index} in final output")
            done.add(source_index)
    return done


def open_state(path: Path, input_path: Path, tolerance: float) -> sqlite3.Connection:
    """Open and initialize the durable per-strategy checkpoint database."""
    connection = sqlite3.connect(path)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=FULL")
    connection.execute(
        "CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL)"
    )
    connection.execute(
        """CREATE TABLE IF NOT EXISTS candidates (
               source_index INTEGER NOT NULL,
               strategy_index INTEGER NOT NULL,
               violated_requirement TEXT NOT NULL,
               text TEXT NOT NULL,
               PRIMARY KEY (source_index, strategy_index)
           )"""
    )
    expected = {
        "prompt_version": PROMPT_VERSION,
        "input_path": str(input_path.resolve()),
        "length_tolerance": str(tolerance),
    }
    existing = dict(connection.execute("SELECT key, value FROM metadata"))
    if existing:
        mismatches = [key for key, value in expected.items() if existing.get(key) != value]
        if mismatches:
            connection.close()
            raise ValueError(
                "checkpoint settings do not match this command for: " + ", ".join(mismatches)
            )
    else:
        connection.executemany(
            "INSERT INTO metadata(key, value) VALUES (?, ?)", expected.items()
        )
        connection.commit()
    return connection


def saved_candidates(connection: sqlite3.Connection, source_index: int) -> dict[int, dict[str, str]]:
    """Load all completed strategies for one source row."""
    rows = connection.execute(
        "SELECT strategy_index, violated_requirement, text "
        "FROM candidates WHERE source_index = ?",
        (source_index,),
    )
    return {
        int(index): {"violated_requirement": violated, "text": text}
        for index, violated, text in rows
    }


def save_candidate(
    connection: sqlite3.Connection,
    source_index: int,
    strategy_index: int,
    candidate: dict[str, str],
) -> None:
    """Atomically checkpoint one successful model response."""
    with connection:
        connection.execute(
            "INSERT OR REPLACE INTO candidates "
            "(source_index, strategy_index, violated_requirement, text) VALUES (?, ?, ?, ?)",
            (
                source_index,
                strategy_index,
                candidate["violated_requirement"],
                candidate["text"],
            ),
        )


def delete_candidate(
    connection: sqlite3.Connection, source_index: int, strategy_index: int
) -> None:
    """Remove a bad checkpoint so that the strategy can be regenerated."""
    with connection:
        connection.execute(
            "DELETE FROM candidates WHERE source_index = ? AND strategy_index = ?",
            (source_index, strategy_index),
        )


def delete_row_candidates(connection: sqlite3.Connection, source_index: int) -> None:
    """Discard checkpoints after their row is durably present in the final output."""
    with connection:
        connection.execute(
            "DELETE FROM candidates WHERE source_index = ?", (source_index,)
        )


def input_rows(
    path: Path, completed: set[int], limit: int | None
) -> Iterator[SourceRow]:
    """Stream unfinished non-empty rows from the large input CSV."""
    yielded = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required = {"question", "content_text"}
        if not required.issubset(reader.fieldnames or []):
            raise ValueError(f"CSV must contain columns {sorted(required)}")
        for source_index, item in enumerate(reader):
            if source_index in completed:
                continue
            query = (item["question"] or "").strip()
            positive = (item["content_text"] or "").strip()
            if not query or not positive:
                continue
            yield SourceRow(source_index, query, positive)
            yielded += 1
            if limit is not None and yielded >= limit:
                return


def chunks(rows: Iterator[SourceRow], size: int) -> Iterator[list[SourceRow]]:
    """Yield bounded row batches without loading the source dataset into memory."""
    while True:
        batch: list[SourceRow] = []
        try:
            for _ in range(size):
                batch.append(next(rows))
        except StopIteration:
            pass
        if not batch:
            return
        yield batch


def build_record(row: SourceRow, candidates: dict[int, dict[str, str]]) -> dict[str, Any]:
    """Build one final record in the format consumed by the cleaning scripts."""
    if set(candidates) != set(range(len(STRATEGIES))):
        raise ValueError("cannot finalize a row before all seven strategies are complete")
    negatives = []
    seen: set[str] = set()
    for index, strategy in enumerate(STRATEGIES):
        candidate = candidates[index]
        normalized = " ".join(candidate["text"].split())
        if normalized in seen:
            raise ValueError(f"duplicate text generated for strategy index {index}")
        seen.add(normalized)
        negatives.append(
            {
                "id": index + 1,
                "type": strategy["type"],
                "hardness": strategy["hardness"],
                "violated_requirement": candidate["violated_requirement"],
                "text": candidate["text"],
            }
        )
    return {
        "source_index": row.source_index,
        "query": row.query,
        "positive_document": row.positive,
        "hard_negatives": negatives,
    }


def append_line(handle: Any, value: dict[str, Any]) -> None:
    """Durably append one completed dataset record."""
    handle.write(json.dumps(value, ensure_ascii=False) + "\n")
    handle.flush()
    os.fsync(handle.fileno())


def process_batch(
    batch: list[SourceRow],
    connection: sqlite3.Connection,
    output: Any,
    errors: Any,
    executor: ThreadPoolExecutor,
    args: argparse.Namespace,
) -> tuple[int, int, int]:
    """Generate missing strategies and finalize all possible rows in one batch."""
    futures: dict[Future[dict[str, str]], tuple[SourceRow, int]] = {}
    requested = 0
    for row in batch:
        existing = saved_candidates(connection, row.source_index)
        for strategy_index in range(len(STRATEGIES)):
            if strategy_index not in existing:
                future = executor.submit(generate_candidate, row, strategy_index, args)
                futures[future] = (row, strategy_index)
                requested += 1

    failed_strategies = 0
    for future in as_completed(futures):
        row, strategy_index = futures[future]
        try:
            candidate = future.result()
            save_candidate(connection, row.source_index, strategy_index, candidate)
        except Exception as exc:
            append_line(
                errors,
                {
                    "source_index": row.source_index,
                    "strategy": STRATEGIES[strategy_index]["type"],
                    "error": str(exc),
                    "time": time.time(),
                },
            )
            failed_strategies += 1

    saved_rows = incomplete_rows = 0
    for row in batch:
        candidates = saved_candidates(connection, row.source_index)
        if len(candidates) != len(STRATEGIES):
            incomplete_rows += 1
            continue
        try:
            record = build_record(row, candidates)
        except ValueError as exc:
            # Exact duplicates are unlikely, but leaving them checkpointed would make
            # the row permanently unfinishable. Drop only the later duplicate.
            match = re.search(r"strategy index (\d+)", str(exc))
            if match:
                delete_candidate(connection, row.source_index, int(match.group(1)))
            append_line(
                errors,
                {"source_index": row.source_index, "error": str(exc), "time": time.time()},
            )
            incomplete_rows += 1
            continue
        append_line(output, record)
        # append_line() flushes and fsyncs first, so the final JSONL is authoritative
        # even if the process is interrupted immediately after this cleanup.
        delete_row_candidates(connection, row.source_index)
        saved_rows += 1

    return saved_rows, incomplete_rows, requested


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate resumable hard negatives using seven independent, strategy-specific "
            "OpenAI-compatible requests per source row."
        )
    )
    parser.add_argument("--input", type=Path, default=Path("data/raw/porseman_clean.csv"))
    parser.add_argument(
        "--output", type=Path, default=Path("data/raw/hard_negatives_all_separate.jsonl")
    )
    parser.add_argument(
        "--state",
        type=Path,
        help="SQLite checkpoint path (default: OUTPUT.state.sqlite3).",
    )
    parser.add_argument("--base-url", default="http://127.0.0.1:18000/v1")
    parser.add_argument("--api-key", default="dummy")
    parser.add_argument("--model", default="Qwen/Qwen3.5-397B-A17B-FP8")
    parser.add_argument("--batch-size", type=int, default=50, help="Source rows per batch.")
    parser.add_argument("--workers", type=int, default=50)
    parser.add_argument("--retries", type=int, default=4)
    parser.add_argument("--retry-delay", type=float, default=2.0)
    parser.add_argument("--timeout", type=float, default=300.0)
    parser.add_argument("--temperature", type=float, default=0.3)
    parser.add_argument("--max-tokens", type=int, default=8192)
    parser.add_argument(
        "--length-tolerance",
        type=float,
        default=0.15,
        help=(
            "Accepted fractional word-count difference from the positive passage; "
            "0.15 means +/-15%% (default: 0.15)."
        ),
    )
    thinking = parser.add_mutually_exclusive_group()
    thinking.add_argument("--thinking", dest="enable_thinking", action="store_true")
    thinking.add_argument("--no-thinking", dest="enable_thinking", action="store_false")
    parser.set_defaults(enable_thinking=True)
    parser.add_argument(
        "--limit", type=int, help="Process at most this many unfinished source rows."
    )
    args = parser.parse_args()

    if min(args.batch_size, args.workers, args.retries) < 1:
        parser.error("batch-size, workers, and retries must be positive")
    if not 0 <= args.length_tolerance <= 1:
        parser.error("--length-tolerance must be between 0 and 1")
    if args.timeout <= 0 or args.retry_delay < 0 or args.max_tokens < 1:
        parser.error("timeout and max-tokens must be positive; retry-delay cannot be negative")
    if not args.input.is_file():
        parser.error(f"input file not found: {args.input}")
    if args.input.resolve() == args.output.resolve():
        parser.error("output path must be different from input")
    if args.state is None:
        args.state = args.output.with_suffix(args.output.suffix + ".state.sqlite3")
    return args


def main() -> None:
    args = parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.state.parent.mkdir(parents=True, exist_ok=True)
    repair_output_tail(args.output)
    completed = completed_indices(args.output)
    error_path = args.output.with_suffix(args.output.suffix + ".errors.jsonl")
    connection = open_state(args.state, args.input, args.length_tolerance)
    rows = input_rows(args.input, completed, args.limit)
    saved_total = incomplete_total = requested_total = processed_total = 0

    print(
        f"Resuming with {len(completed)} completed rows; output={args.output}; "
        f"checkpoint={args.state}"
    )
    try:
        with args.output.open("a", encoding="utf-8", newline="\n") as output, \
                error_path.open("a", encoding="utf-8", newline="\n") as errors, \
                ThreadPoolExecutor(max_workers=args.workers) as executor, \
                tqdm(unit="row", dynamic_ncols=True, desc="Hard-negative rows") as progress:
            for batch in chunks(rows, args.batch_size):
                saved, incomplete, requested = process_batch(
                    batch, connection, output, errors, executor, args
                )
                saved_total += saved
                incomplete_total += incomplete
                requested_total += requested
                processed_total += len(batch)
                progress.update(len(batch))
                progress.set_postfix(
                    saved=saved_total,
                    incomplete=incomplete_total,
                    requests=requested_total,
                    refresh=False,
                )
    except KeyboardInterrupt:
        print("Interrupted safely. Run the identical command to resume saved strategies.")
        raise SystemExit(130)
    finally:
        connection.close()

    print(
        f"Finished this run: rows_seen={processed_total}, saved={saved_total}, "
        f"incomplete={incomplete_total}, model_requests={requested_total}."
    )
    if incomplete_total:
        print("Some rows remain incomplete. Run the identical command again to retry only missing strategies.")
        raise SystemExit(1)
    print(f"Dataset is complete for all selected unfinished rows: {args.output}")


if __name__ == "__main__":
    main()
