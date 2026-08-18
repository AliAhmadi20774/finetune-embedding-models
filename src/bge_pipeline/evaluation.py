from __future__ import annotations

import math
import random
from collections import defaultdict
from typing import Any

import numpy as np


def build_retrieval_data(rows: list[dict[str, Any]]):
    corpus: dict[str, str] = {}
    queries: dict[str, str] = {}
    relevant: dict[str, set[str]] = defaultdict(set)
    for row in rows:
        corpus[row["document_id"]] = row["content_text"]
        queries[row["question_id"]] = row["question"]
        relevant[row["question_id"]].add(row["document_id"])
    return queries, corpus, dict(relevant)


def retrieval_metrics(model, rows: list[dict[str, Any]], batch_size: int = 1) -> dict[str, float]:
    queries, corpus, relevant = build_retrieval_data(rows)
    query_ids, corpus_ids = list(queries), list(corpus)
    query_embeddings = model.encode(
        [queries[key] for key in query_ids], batch_size=batch_size, normalize_embeddings=True,
        convert_to_numpy=True, show_progress_bar=True,
    )
    corpus_embeddings = model.encode(
        [corpus[key] for key in corpus_ids], batch_size=batch_size, normalize_embeddings=True,
        convert_to_numpy=True, show_progress_bar=True,
    )
    if hasattr(query_embeddings, "detach"):
        query_embeddings = query_embeddings.detach().cpu().numpy()
    if hasattr(corpus_embeddings, "detach"):
        corpus_embeddings = corpus_embeddings.detach().cpu().numpy()
    top_k = min(10, len(corpus_ids))
    ranked_hits: list[list[int]] = []
    # Chunking avoids materializing the full query-by-corpus score matrix.
    for start in range(0, len(query_embeddings), 128):
        scores = np.asarray(query_embeddings[start:start + 128]) @ np.asarray(corpus_embeddings).T
        candidate_ids = np.argpartition(-scores, kth=top_k - 1, axis=1)[:, :top_k]
        for row_scores, candidates in zip(scores, candidate_ids):
            ranked_hits.append(candidates[np.argsort(-row_scores[candidates])].tolist())
    recall = {1: 0, 5: 0, 10: 0}
    reciprocal_ranks: list[float] = []
    ndcgs: list[float] = []
    for query_id, query_hits in zip(query_ids, ranked_hits):
        positives = relevant[query_id]
        ranked = [corpus_ids[index] for index in query_hits]
        for k in recall:
            recall[k] += int(any(document_id in positives for document_id in ranked[:k]))
        positive_ranks = [rank for rank, document_id in enumerate(ranked, 1) if document_id in positives]
        reciprocal_ranks.append(1 / min(positive_ranks) if positive_ranks else 0.0)
        dcg = sum(1 / math.log2(rank + 1) for rank in positive_ranks)
        ideal_count = min(len(positives), 10)
        idcg = sum(1 / math.log2(rank + 1) for rank in range(1, ideal_count + 1))
        ndcgs.append(dcg / idcg if idcg else 0.0)
    count = len(query_ids)
    return {
        "queries": count,
        "corpus_documents": len(corpus_ids),
        "recall@1": recall[1] / count,
        "recall@5": recall[5] / count,
        "recall@10": recall[10] / count,
        "mrr@10": float(np.mean(reciprocal_ranks)),
        "ndcg@10": float(np.mean(ndcgs)),
    }


def pair_diagnostics(model, rows: list[dict[str, Any]], batch_size: int = 1, seed: int = 42) -> dict[str, float]:
    questions = [row["question"] for row in rows]
    answers = [row["content_text"] for row in rows]
    q_embeddings = model.encode(questions, batch_size=batch_size, normalize_embeddings=True, show_progress_bar=True)
    a_embeddings = model.encode(answers, batch_size=batch_size, normalize_embeddings=True, show_progress_bar=True)
    positive = np.sum(q_embeddings * a_embeddings, axis=1)
    indices = list(range(len(rows)))
    random.Random(seed).shuffle(indices)
    if len(indices) > 1 and any(indices[i] == i for i in range(len(indices))):
        indices = indices[1:] + indices[:1]
    negative = np.sum(q_embeddings * a_embeddings[indices], axis=1)
    return {
        "positive_cosine_mean": float(np.mean(positive)),
        "positive_cosine_median": float(np.median(positive)),
        "random_negative_cosine_mean": float(np.mean(negative)),
        "random_negative_cosine_median": float(np.median(negative)),
        "mean_margin": float(np.mean(positive - negative)),
    }
