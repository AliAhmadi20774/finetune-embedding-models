import csv
import json
from pathlib import Path

import pytest

from bge_pipeline.data import export_duplicate_question_extras, normalize_persian, read_and_split, stable_id


def write_csv(path: Path, rows):
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["question", "content_text"])
        writer.writeheader()
        writer.writerows(rows)


def test_persian_normalization_is_conservative():
    assert normalize_persian("  يكي\nاز   كلمه‌ها؟ ") == "یکی از کلمه‌ها؟"
    assert normalize_persian("سؤال: ۲ + ۲؟") == "سؤال: ۲ + ۲؟"


def test_stable_id_is_deterministic():
    assert stable_id("الف", "ب") == stable_id("الف", "ب")
    assert stable_id("الف", "ب") != stable_id("ب", "الف")


def test_dedup_and_grouped_split_are_deterministic(tmp_path):
    source = tmp_path / "data.csv"
    rows = []
    for index in range(30):
        rows.append({"question": f"سؤال {index}", "content_text": f"پاسخ {index}"})
    rows += [rows[0], {"question": "سؤال 0", "content_text": "پاسخ دوم"}]
    write_csv(source, rows)
    first, report = read_and_split(source, seed=42)
    second, _ = read_and_split(source, seed=42)
    assert [(r["id"], r["split"]) for r in first] == [(r["id"], r["split"]) for r in second]
    assert report["duplicate_pair_rows_removed"] == 1
    question_splits = {r["split"] for r in first if r["normalized_question"] == "سؤال 0"}
    assert len(question_splits) == 1
    assert set(report["split_counts"]) == {"train", "validation", "test"}


def test_schema_and_empty_values_are_rejected(tmp_path):
    wrong = tmp_path / "wrong.csv"
    wrong.write_text("q,a\nx,y\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Expected CSV columns"):
        read_and_split(wrong)
    empty = tmp_path / "empty.csv"
    write_csv(empty, [{"question": "", "content_text": "پاسخ"}])
    with pytest.raises(ValueError, match="Empty question"):
        read_and_split(empty)


def test_duplicate_extra_export(tmp_path):
    source = tmp_path / "data.csv"
    output = tmp_path / "duplicates.csv"
    write_csv(source, [
        {"question": "سوال", "content_text": "پاسخ یک"},
        {"question": "سوال", "content_text": "پاسخ دو"},
        {"question": "سوال", "content_text": "پاسخ یک"},
    ])
    report = export_duplicate_question_extras(source, output)
    assert report == {"duplicate_question_extra_rows": 2, "exact_pair_duplicate_extra_rows": 1}
    with output.open("r", encoding="utf-8-sig", newline="") as handle:
        exported = list(csv.DictReader(handle))
    assert len(exported) == 2
    assert exported[-1]["is_exact_pair_duplicate"] == "True"
