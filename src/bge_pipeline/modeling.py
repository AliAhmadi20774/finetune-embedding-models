from __future__ import annotations

import json
from pathlib import Path
from typing import Any

MODEL_NAME = "BAAI/bge-m3"
MAX_LENGTH = 8192


def load_model(model_name: str = MODEL_NAME, device: str | None = None):
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer(model_name, device=device)
    model.max_seq_length = MAX_LENGTH
    return model


def inspect_model(model) -> dict[str, Any]:
    transformer = model[0]
    config = transformer.auto_model.config
    modules = [type(module).__name__ for module in model]
    total = sum(parameter.numel() for parameter in model.parameters())
    trainable = sum(parameter.numel() for parameter in model.parameters() if parameter.requires_grad)
    pooling = model[1]
    details = {
        "model_type": config.model_type,
        "architecture": list(getattr(config, "architectures", [])),
        "hidden_size": config.hidden_size,
        "num_hidden_layers": config.num_hidden_layers,
        "num_attention_heads": config.num_attention_heads,
        "max_position_embeddings": config.max_position_embeddings,
        "sentence_transformer_max_seq_length": model.max_seq_length,
        "embedding_dimension": model.get_sentence_embedding_dimension(),
        "vocab_size": config.vocab_size,
        "modules": modules,
        "pooling": {
            "cls": bool(pooling.pooling_mode_cls_token),
            "mean": bool(pooling.pooling_mode_mean_tokens),
            "max": bool(pooling.pooling_mode_max_tokens),
        },
        "has_normalize_module": "Normalize" in modules,
        "total_parameters": total,
        "trainable_parameters": trainable,
        "dtype": str(next(model.parameters()).dtype),
        "tokenizer_class": type(model.tokenizer).__name__,
    }
    expected = {
        "model_type": "xlm-roberta",
        "hidden_size": 1024,
        "num_hidden_layers": 24,
        "num_attention_heads": 16,
        "max_position_embeddings": 8194,
        "embedding_dimension": 1024,
    }
    failures = {key: {"expected": value, "actual": details[key]} for key, value in expected.items() if details[key] != value}
    if not details["pooling"]["cls"] or details["pooling"]["mean"]:
        failures["pooling"] = {"expected": "CLS only", "actual": details["pooling"]}
    if not details["has_normalize_module"]:
        failures["normalize"] = {"expected": True, "actual": False}
    details["validation_failures"] = failures
    return details


def write_json(payload: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
