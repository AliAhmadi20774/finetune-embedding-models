"""Find semantically near-duplicate questions across Porseman train and test.

The dense BGE-M3 vectors are L2-normalized before comparison, therefore their
dot product is cosine similarity.  Every train/test pair at or above the
threshold is written to a CSV file; a second CSV contains the best train match
for *every* test question (including questions below the threshold).
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path
import re
import unicodedata

import numpy as np


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_TRAIN = PROJECT_ROOT / "data" / "processed" / "porseman_train.csv"
DEFAULT_TEST = PROJECT_ROOT / "data" / "processed" / "porseman_test.csv"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "porseman_train_test_similarity"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find test questions whose cosine similarity to train is high."
    )
    parser.add_argument("--train", type=Path, default=DEFAULT_TRAIN)
    parser.add_argument("--test", type=Path, default=DEFAULT_TEST)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--model", default="BAAI/bge-m3", help="HF model name or local path.")
    parser.add_argument("--threshold", type=float, default=0.90)
    parser.add_argument("--device", default=None, help="For example: cuda:0 or cpu.")
    parser.add_argument("--batch-size", type=int, default=32)
    parser.add_argument("--max-length", type=int, default=512)
    parser.add_argument(
        "--train-chunk-size",
        type=int,
        default=4096,
        help="Number of train embeddings compared at once (limits RAM/VRAM use).",
    )
    parser.add_argument("--no-fp16", action="store_true", help="Disable FP16 on GPU.")
    return parser.parse_args()


def read_questions(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise FileNotFoundError(f"CSV file not found: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        rows = list(csv.DictReader(file))
    if not rows or "id" not in rows[0] or "question" not in rows[0]:
        raise ValueError(f"{path} must contain id and question columns.")
    valid = [row for row in rows if row["question"].strip()]
    print(f"Loaded {len(valid):,} non-empty questions from {path.name}")
    return valid


def normalize_rows(vectors: np.ndarray) -> np.ndarray:
    vectors = np.asarray(vectors, dtype=np.float32)
    norms = np.linalg.norm(vectors, axis=1, keepdims=True)
    return vectors / np.maximum(norms, 1e-12)


def normalize_question_for_exact_match(question: str) -> str:
    """Normalize harmless Persian/Arabic spelling differences for exact-match checks."""
    question = unicodedata.normalize("NFKC", question)
    question = question.translate(str.maketrans({"ي": "ی", "ك": "ک", "ى": "ی"}))
    return re.sub(r"[\W_]+", "", question, flags=re.UNICODE).casefold()


def encode_questions(model, rows: list[dict[str, str]], args: argparse.Namespace) -> np.ndarray:
    # encode_queries uses BGE-M3's query encoding path and returns dense vectors.
    embeddings = model.encode_queries(
        [row["question"] for row in rows],
        batch_size=args.batch_size,
        max_length=args.max_length,
        return_dense=True,
        return_sparse=False,
        return_colbert_vecs=False,
    )["dense_vecs"]
    return normalize_rows(embeddings)


def main() -> None:
    args = parse_args()
    if not 0 <= args.threshold <= 1:
        raise ValueError("--threshold must be between 0 and 1.")
    if min(args.batch_size, args.max_length, args.train_chunk_size) <= 0:
        raise ValueError("Batch size, max length, and chunk size must be positive.")

    try:
        from FlagEmbedding import BGEM3FlagModel
    except ImportError as error:
        raise RuntimeError("Install dependencies first: pip install -r requirements.txt") from error

    train_rows = read_questions(args.train)
    test_rows = read_questions(args.test)
    model_kwargs = {"use_fp16": not args.no_fp16, "pooling_method": "cls"}
    if args.device:
        model_kwargs["devices"] = [args.device]
    model = BGEM3FlagModel(args.model, **model_kwargs)

    print("Encoding train questions …")
    train_vectors = encode_questions(model, train_rows, args)
    print("Encoding test questions …")
    test_vectors = encode_questions(model, test_rows, args)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    pairs_path = args.output_dir / "test_train_cosine_at_or_above_threshold.csv"
    best_path = args.output_dir / "best_train_match_for_each_test_question.csv"
    match_count = 0
    suspicious_test_count = 0
    exact_match_count = 0
    train_normalized_questions = {
        normalize_question_for_exact_match(row["question"]) for row in train_rows
    }

    with pairs_path.open("w", encoding="utf-8-sig", newline="") as pairs_file, best_path.open(
        "w", encoding="utf-8-sig", newline=""
    ) as best_file:
        pairs = csv.DictWriter(
            pairs_file,
            fieldnames=[
                "test_id", "test_question", "train_id", "train_question",
                "cosine_similarity", "cosine_distance",
            ],
        )
        best = csv.DictWriter(
            best_file,
            fieldnames=[
                "test_id", "test_question", "best_train_id", "best_train_question",
                "cosine_similarity", "cosine_distance", "meets_threshold", "exact_text_match",
            ],
        )
        pairs.writeheader()
        best.writeheader()

        for test_index, test_row in enumerate(test_rows):
            query = test_vectors[test_index]
            best_index, best_score = -1, -np.inf
            for start in range(0, len(train_rows), args.train_chunk_size):
                stop = min(start + args.train_chunk_size, len(train_rows))
                # Unit vectors: dot product == cosine similarity.
                scores = train_vectors[start:stop] @ query
                local_best = int(np.argmax(scores))
                if scores[local_best] > best_score:
                    best_score = float(scores[local_best])
                    best_index = start + local_best
                for local_index in np.flatnonzero(scores >= args.threshold):
                    train_index = start + int(local_index)
                    score = float(scores[local_index])
                    pairs.writerow({
                        "test_id": test_row["id"], "test_question": test_row["question"],
                        "train_id": train_rows[train_index]["id"],
                        "train_question": train_rows[train_index]["question"],
                        "cosine_similarity": f"{score:.6f}",
                        "cosine_distance": f"{1 - score:.6f}",
                    })
                    match_count += 1
            train_row = train_rows[best_index]
            meets_threshold = best_score >= args.threshold
            exact_text_match = (
                normalize_question_for_exact_match(test_row["question"])
                in train_normalized_questions
            )
            best.writerow({
                "test_id": test_row["id"], "test_question": test_row["question"],
                "best_train_id": train_row["id"], "best_train_question": train_row["question"],
                "cosine_similarity": f"{best_score:.6f}",
                "cosine_distance": f"{1 - best_score:.6f}",
                "meets_threshold": meets_threshold,
                "exact_text_match": exact_text_match,
            })
            suspicious_test_count += int(meets_threshold)
            exact_match_count += int(exact_text_match)
            if (test_index + 1) % 100 == 0 or test_index + 1 == len(test_rows):
                print(f"Compared {test_index + 1:,}/{len(test_rows):,} test questions")

    print(f"Found {match_count:,} train/test pairs with cosine >= {args.threshold:.3f}")
    print(
        f"Potentially leaked test questions: {suspicious_test_count:,}/{len(test_rows):,} "
        f"({suspicious_test_count / len(test_rows):.2%})"
    )
    print(f"Exact text matches after Persian normalization: {exact_match_count:,}")
    print(f"All matching pairs: {pairs_path}")
    print(f"Best match per test question: {best_path}")


if __name__ == "__main__":
    main()
