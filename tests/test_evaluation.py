import numpy as np

from bge_pipeline.evaluation import build_retrieval_data, pair_diagnostics, retrieval_metrics


class TinyModel:
    vectors = {
        "q1": np.array([1.0, 0.0], dtype=np.float32),
        "a1": np.array([1.0, 0.0], dtype=np.float32),
        "q2": np.array([0.0, 1.0], dtype=np.float32),
        "a2": np.array([0.0, 1.0], dtype=np.float32),
    }

    def encode(self, texts, **kwargs):
        values = np.stack([self.vectors[text] for text in texts])
        if kwargs.get("convert_to_tensor"):
            import torch
            return torch.from_numpy(values)
        return values


ROWS = [
    {"question_id": "qid1", "document_id": "did1", "question": "q1", "content_text": "a1"},
    {"question_id": "qid2", "document_id": "did2", "question": "q2", "content_text": "a2"},
]


def test_retrieval_metrics_perfect_ranking():
    metrics = retrieval_metrics(TinyModel(), ROWS)
    assert metrics["recall@1"] == 1.0
    assert metrics["mrr@10"] == 1.0
    assert metrics["ndcg@10"] == 1.0


def test_pair_diagnostics_margin():
    result = pair_diagnostics(TinyModel(), ROWS)
    assert result["positive_cosine_mean"] == 1.0
    assert result["random_negative_cosine_mean"] == 0.0
    assert result["mean_margin"] == 1.0
