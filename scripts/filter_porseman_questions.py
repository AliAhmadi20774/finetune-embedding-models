"""Filter Porseman CSV rows to questions with a terminal question mark.

The script also removes rows whose ``question`` field looks like an article or
an answer accidentally placed in the question column.  It prints a detailed
summary so the filtering decision is reproducible and auditable.
"""

from __future__ import annotations

import argparse
import csv
import re
import unicodedata
from collections import Counter
from pathlib import Path
from typing import TextIO


QUESTION_MARKS = ("؟", "?")
ARTICLE_MARKERS = (
    "در این مقاله",
    "در این نوشتار",
    "در شماره قبل",
    "در شماره های قبل",
    "در شماره‌های قبل",
    "ادامه دارد",
    "پی نوشت",
    "پی‌نوشت",
    "کلید واژه",
    "کلیدواژه",
    "فهرست منابع",
)


def normalize(text: str) -> str:
    """Normalize text for exact cross-column comparison."""
    text = unicodedata.normalize("NFKC", text or "")
    text = text.translate(
        str.maketrans({"ي": "ی", "ى": "ی", "ك": "ک", "ؤ": "و", "ۀ": "ه", "ة": "ه"})
    )
    text = re.sub(r"[\u200c\u200d\ufeff\u00ad]", " ", text)
    return re.sub(r"\s+", " ", text).strip().lower()


def has_required_question_mark(question: str, position: str) -> bool:
    """Return whether question contains a mark at the requested position."""
    stripped = question.rstrip()
    if position == "end":
        return stripped.endswith(QUESTION_MARKS)
    return any(mark in question for mark in QUESTION_MARKS)


def suspicious_reasons(question: str, answer: str, normalized_answers: set[str]) -> list[str]:
    """Return reasons for classifying a question as likely article/answer text."""
    normalized_question = normalize(question)
    if normalized_question in normalized_answers:
        return ["question_matches_an_answer"]

    question_length = len(normalized_question)
    answer_length = len(normalize(answer))
    if question_length >= 1_200:
        return ["very_long_question"]

    signals: list[str] = []
    if question_length >= 500:
        signals.append("long_question")
    if question_length >= 300 and question_length >= max(1, answer_length) * 2:
        signals.append("question_much_longer_than_answer")
    if any(marker in normalized_question for marker in ARTICLE_MARKERS):
        signals.append("article_phrase")
    if len(re.findall(r"\[\s*[۰-۹0-9]+\s*\]", normalized_question)) >= 3:
        signals.append("many_citations")
    if len(re.findall(r"(?:^|\s)[۰-۹0-9]+[.)ـ-]", normalized_question)) >= 4:
        signals.append("numbered_list")
    if question.count("\n") >= 3:
        signals.append("many_paragraphs")

    # Requiring three independent signals keeps ordinary long user questions.
    return signals if len(signals) >= 3 else []


def collect_normalized_answers(input_path: Path) -> set[str]:
    answers: set[str] = set()
    with input_path.open("r", encoding="utf-8-sig", newline="") as source:
        reader = csv.DictReader(source)
        validate_fields(reader)
        for row in reader:
            answer = normalize(row["content_text"])
            if answer:
                answers.add(answer)
    return answers


def validate_fields(reader: csv.DictReader) -> None:
    required = {"question", "content_text"}
    missing = required.difference(reader.fieldnames or ())
    if missing:
        raise ValueError(f"Missing required CSV columns: {', '.join(sorted(missing))}")


def filter_csv(
    input_path: Path,
    output_path: Path | None,
    rejected_path: Path | None,
    question_mark_position: str,
) -> Counter[str]:
    normalized_answers = collect_normalized_answers(input_path)
    stats: Counter[str] = Counter()

    output_file: TextIO | None = None
    rejected_file: TextIO | None = None
    try:
        if output_path:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_file = output_path.open("w", encoding="utf-8-sig", newline="")
        if rejected_path:
            rejected_path.parent.mkdir(parents=True, exist_ok=True)
            rejected_file = rejected_path.open("w", encoding="utf-8-sig", newline="")

        with input_path.open("r", encoding="utf-8-sig", newline="") as source:
            reader = csv.DictReader(source)
            validate_fields(reader)
            fieldnames = list(reader.fieldnames or ())
            writer = csv.DictWriter(output_file, fieldnames=fieldnames) if output_file else None
            rejected_writer = (
                csv.DictWriter(rejected_file, fieldnames=[*fieldnames, "rejection_reason"])
                if rejected_file
                else None
            )
            if writer:
                writer.writeheader()
            if rejected_writer:
                rejected_writer.writeheader()

            for row in reader:
                stats["total_rows"] += 1
                question = row["question"] or ""
                if not has_required_question_mark(question, question_mark_position):
                    reasons = ["missing_question_mark"]
                else:
                    reasons = suspicious_reasons(
                        question, row["content_text"] or "", normalized_answers
                    )

                if reasons:
                    stats["removed_rows"] += 1
                    for reason in reasons:
                        stats[f"removed:{reason}"] += 1
                    if rejected_writer:
                        rejected_writer.writerow({**row, "rejection_reason": ";".join(reasons)})
                    continue

                stats["kept_rows"] += 1
                if writer:
                    writer.writerow(row)
    finally:
        if output_file:
            output_file.close()
        if rejected_file:
            rejected_file.close()

    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, required=True, help="Input CSV path")
    parser.add_argument("--output", type=Path, help="Filtered CSV path")
    parser.add_argument(
        "--rejected-output", type=Path, help="Optional CSV containing removed rows and reasons"
    )
    parser.add_argument(
        "--question-mark-position",
        choices=("end", "any"),
        default="end",
        help="Require a question mark at the end (default) or anywhere in the question",
    )
    parser.add_argument(
        "--dry-run", action="store_true", help="Only calculate statistics; do not write files"
    )
    args = parser.parse_args()
    if not args.dry_run and args.output is None:
        parser.error("--output is required unless --dry-run is used")
    return args


def main() -> None:
    args = parse_args()
    stats = filter_csv(
        input_path=args.input,
        output_path=None if args.dry_run else args.output,
        rejected_path=None if args.dry_run else args.rejected_output,
        question_mark_position=args.question_mark_position,
    )
    for key in sorted(stats):
        print(f"{key}: {stats[key]}")


if __name__ == "__main__":
    main()
