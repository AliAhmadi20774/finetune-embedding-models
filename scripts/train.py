from __future__ import annotations

import argparse
import json
import shutil
import sys
import traceback
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bge_pipeline.data import load_jsonl
from bge_pipeline.evaluation import build_retrieval_data
from bge_pipeline.modeling import MAX_LENGTH, MODEL_NAME, inspect_model, load_model, write_json


def parse_args():
    parser = argparse.ArgumentParser(description="Full dense fine-tuning of BGE-M3 with a strict 8,192-token limit.")
    parser.add_argument("--model", default=MODEL_NAME)
    parser.add_argument("--train-file", type=Path, default=Path("data/processed/train.jsonl"))
    parser.add_argument("--validation-file", type=Path, default=Path("data/processed/validation.jsonl"))
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/bge-m3-dense"))
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--epochs", type=float, default=1.0)
    parser.add_argument("--learning-rate", type=float, default=1e-5)
    parser.add_argument("--outer-batch-size", type=int, default=8)
    parser.add_argument("--mini-batch-size", type=int, default=1)
    parser.add_argument("--save-steps", type=int, default=500)
    parser.add_argument("--logging-steps", type=int, default=10)
    parser.add_argument("--smoke-test", action="store_true", help="Run exactly two optimizer steps and verify reload.")
    parser.add_argument("--resume-from-checkpoint", default=None, help="Checkpoint path, or 'latest' to resume the newest checkpoint.")
    return parser.parse_args()


def memory_payload(args, status: str, error: str | None = None) -> dict:
    import torch

    cuda = torch.cuda.is_available()
    return {
        "status": status,
        "error": error,
        "model": args.model,
        "max_sequence_length": MAX_LENGTH,
        "full_fine_tuning": True,
        "outer_batch_size": args.outer_batch_size,
        "cached_loss_mini_batch_size": args.mini_batch_size,
        "fp16": True,
        "gradient_checkpointing": True,
        "optimizer": "adafactor",
        "gpu": torch.cuda.get_device_name(0) if cuda else None,
        "peak_allocated_bytes": torch.cuda.max_memory_allocated() if cuda else None,
        "peak_reserved_bytes": torch.cuda.max_memory_reserved() if cuda else None,
    }


def main() -> None:
    args = parse_args()
    if not args.train_file.exists():
        raise SystemExit(f"Missing {args.train_file}; run scripts/prepare_data.py first")

    import torch
    from datasets import Dataset, load_dataset
    from sentence_transformers import SentenceTransformerTrainer, SentenceTransformerTrainingArguments, losses
    from sentence_transformers.evaluation import InformationRetrievalEvaluator

    if not torch.cuda.is_available():
        raise SystemExit("CUDA is required for this explicitly FP16 full fine-tuning run")
    torch.cuda.reset_peak_memory_stats()
    output_dir = args.output_dir / ("smoke" if args.smoke_test else "run")
    output_dir.mkdir(parents=True, exist_ok=True)
    serializable_args = {key: str(value) if isinstance(value, Path) else value for key, value in vars(args).items()}
    write_json(serializable_args, output_dir / "run_config.json")

    try:
        model = load_model(args.model, device="cuda")
        model.max_seq_length = MAX_LENGTH
        # Required when checkpointing starts from integer token IDs; otherwise
        # PyTorch can skip gradients through the checkpointed transformer block.
        model[0].auto_model.enable_input_require_grads()
        model[0].auto_model.gradient_checkpointing_enable()
        structure = inspect_model(model)
        if structure["validation_failures"]:
            raise RuntimeError(f"Unexpected BGE-M3 structure: {structure['validation_failures']}")
        write_json(structure, output_dir / "model_inspection.json")

        train_dataset = load_dataset("json", data_files=str(args.train_file), split="train")
        train_dataset = train_dataset.select_columns(["question", "content_text"]).rename_columns(
            {"question": "anchor", "content_text": "positive"}
        )
        if args.smoke_test:
            train_dataset = train_dataset.select(range(min(len(train_dataset), max(16, args.outer_batch_size * 2))))
        gather_across_devices = int(__import__("os").environ.get("WORLD_SIZE", "1")) > 1
        loss = losses.CachedMultipleNegativesRankingLoss(
            model=model,
            mini_batch_size=args.mini_batch_size,
            scale=50.0,
            gather_across_devices=gather_across_devices,
        )
        argument_values = dict(
            output_dir=str(output_dir / "checkpoints"),
            num_train_epochs=args.epochs,
            max_steps=2 if args.smoke_test else -1,
            per_device_train_batch_size=args.outer_batch_size,
            learning_rate=args.learning_rate,
            warmup_ratio=0.1,
            fp16=True,
            gradient_checkpointing=True,
            max_grad_norm=1.0,
            optim="adafactor",
            save_strategy="steps" if args.smoke_test else "epoch",
            save_steps=1 if args.smoke_test else args.save_steps,
            save_total_limit=2,
            logging_steps=1 if args.smoke_test else args.logging_steps,
            dataloader_drop_last=True,
            seed=args.seed,
            data_seed=args.seed,
            report_to="none",
        )
        evaluator = None
        validation_dataset = None
        if not args.smoke_test:
            validation_rows = load_jsonl(args.validation_file)
            queries, corpus, relevant = build_retrieval_data(validation_rows)
            evaluator = InformationRetrievalEvaluator(
                queries=queries,
                corpus=corpus,
                relevant_docs=relevant,
                name="validation",
                accuracy_at_k=[1, 5, 10],
                mrr_at_k=[10],
                ndcg_at_k=[10],
                show_progress_bar=True,
                batch_size=1,
                main_score_function="cosine",
            )
            validation_dataset = Dataset.from_list([
                {"anchor": row["question"], "positive": row["content_text"]} for row in validation_rows
            ])
            argument_values.update(
                eval_strategy="epoch",
                load_best_model_at_end=True,
                metric_for_best_model="eval_validation_cosine_mrr@10",
                greater_is_better=True,
            )
        training_args = SentenceTransformerTrainingArguments(**argument_values)
        trainer = SentenceTransformerTrainer(
            model=model,
            args=training_args,
            train_dataset=train_dataset,
            eval_dataset=validation_dataset,
            evaluator=evaluator,
            loss=loss,
        )
        resume = True if args.resume_from_checkpoint == "latest" else args.resume_from_checkpoint
        trainer.train(resume_from_checkpoint=resume)
        final_dir = output_dir / "final"
        if trainer.is_world_process_zero():
            model.save_pretrained(str(final_dir))

            # A reload catches incomplete Sentence Transformers checkpoints.
            reloaded = load_model(str(final_dir), device="cuda")
            probe = reloaded.encode(["آزمون بازیابی"], normalize_embeddings=True, convert_to_numpy=True)
            if probe.shape != (1, 1024):
                raise RuntimeError(f"Reloaded checkpoint returned unexpected shape {probe.shape}")

            manifest_source = args.train_file.parent / "manifest.json"
            if manifest_source.exists():
                shutil.copy2(manifest_source, output_dir / "data_manifest.json")
            write_json(memory_payload(args, "completed"), output_dir / "memory_report.json")
            print(f"Training completed; checkpoint: {final_dir}")
    except Exception as exc:
        message = f"{type(exc).__name__}: {exc}"
        is_oom = isinstance(exc, torch.OutOfMemoryError) or "out of memory" in str(exc).lower()
        rank = int(__import__("os").environ.get("RANK", "0"))
        write_json(memory_payload(args, "cuda_oom" if is_oom else "failed", message), output_dir / f"memory_report_rank{rank}.json")
        (output_dir / f"failure_traceback_rank{rank}.txt").write_text(traceback.format_exc(), encoding="utf-8")
        if is_oom:
            print("CUDA OOM at the requested 8,192 tokens. No automatic fallback was applied.", file=sys.stderr)
        raise


if __name__ == "__main__":
    main()
