"""Fine-tune the dense retrieval representation of BGE-M3 on JSONL triples."""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Launch FlagEmbedding dense-only BGE-M3 fine-tuning on {query, pos, neg} JSONL data."
    )
    parser.add_argument("--train-file", type=Path, default=Path("data/splits/train.jsonl"))
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/bge-m3-dense"))
    parser.add_argument("--model", default="BAAI/bge-m3")
    parser.add_argument("--num-gpus", type=int, default=1, help="Number of local GPU processes.")
    parser.add_argument(
        "--cuda-visible-devices",
        help="Physical GPU index or indices to use, for example '0' or '0,1'.",
    )
    parser.add_argument("--epochs", type=float, default=1.0)
    parser.add_argument("--learning-rate", type=float, default=1e-5)
    parser.add_argument("--batch-size", type=int, default=1, help="Per-GPU query-group batch size.")
    parser.add_argument("--train-group-size", type=int, default=8, help="One positive plus seven negatives.")
    parser.add_argument("--query-max-length", type=int, default=512)
    parser.add_argument("--passage-max-length", type=int, default=2048)
    parser.add_argument("--precision", choices=("fp16", "bf16", "fp32"), default="fp16")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--warmup-ratio", type=float, default=0.1)
    parser.add_argument("--save-steps", type=int, default=1000)
    parser.add_argument("--logging-steps", type=int, default=10)
    parser.add_argument("--save-total-limit", type=int, default=2)
    parser.add_argument("--cache-dir", type=Path, default=Path("cache/models"))
    parser.add_argument("--data-cache-dir", type=Path, default=Path("cache/flagembedding"))
    parser.add_argument(
        "--resume-from-checkpoint",
        type=Path,
        help="Resume optimizer, scheduler, and trainer state from a checkpoint directory.",
    )
    parser.add_argument("--overwrite-output-dir", action="store_true")
    parser.add_argument("--dry-run", action="store_true", help="Print the underlying command without training.")
    return parser.parse_args()


def build_command(args: argparse.Namespace) -> list[str]:
    training_module = (
        "scripts.flagembedding_m3_resume"
        if args.resume_from_checkpoint is not None
        else "FlagEmbedding.finetune.embedder.encoder_only.m3"
    )
    command = [
        sys.executable,
        "-m",
        "torch.distributed.run",
        "--nproc_per_node",
        str(args.num_gpus),
        "-m",
        training_module,
        "--model_name_or_path",
        args.model,
        "--cache_dir",
        str(args.cache_dir),
        "--train_data",
        str(args.train_file),
        "--cache_path",
        str(args.data_cache_dir),
        "--train_group_size",
        str(args.train_group_size),
        "--query_max_len",
        str(args.query_max_length),
        "--passage_max_len",
        str(args.passage_max_length),
        "--pad_to_multiple_of",
        "8",
        "--knowledge_distillation",
        "False",
        "--same_dataset_within_batch",
        "True",
        "--small_threshold",
        "0",
        "--drop_threshold",
        "0",
        "--output_dir",
        str(args.output_dir),
        "--learning_rate",
        str(args.learning_rate),
        "--num_train_epochs",
        str(args.epochs),
        "--per_device_train_batch_size",
        str(args.batch_size),
        "--dataloader_drop_last",
        "True",
        "--warmup_ratio",
        str(args.warmup_ratio),
        "--gradient_checkpointing",
        "--gradient_checkpointing_kwargs",
        '{"use_reentrant": false}',
        "--ddp_find_unused_parameters",
        "True",
        "--logging_steps",
        str(args.logging_steps),
        "--disable_tqdm",
        "False",
        "--save_steps",
        str(args.save_steps),
        "--save_total_limit",
        str(args.save_total_limit),
        "--seed",
        str(args.seed),
        "--temperature",
        "0.02",
        "--sentence_pooling_method",
        "cls",
        "--normalize_embeddings",
        "True",
        "--unified_finetuning",
        "False",
        "--use_self_distill",
        "False",
        "--fix_encoder",
        "False",
    ]
    if args.precision != "fp32":
        command.append(f"--{args.precision}")
    if args.num_gpus > 1:
        command.append("--negatives_cross_device")
    if args.resume_from_checkpoint is not None:
        command.extend(["--resume_from_checkpoint", str(args.resume_from_checkpoint)])
    if args.overwrite_output_dir:
        command.append("--overwrite_output_dir")
    return command


def main() -> None:
    args = parse_args()
    if not args.train_file.is_file():
        raise SystemExit(f"Training file not found: {args.train_file}")
    if args.num_gpus < 1 or args.batch_size < 1 or args.train_group_size < 2:
        raise SystemExit("num-gpus, batch-size, and train-group-size must be positive; group size must be at least 2.")
    if args.cuda_visible_devices:
        selected_gpus = [item.strip() for item in args.cuda_visible_devices.split(",") if item.strip()]
        if len(selected_gpus) != args.num_gpus:
            raise SystemExit("--num-gpus must equal the number of indices in --cuda-visible-devices.")
    if args.resume_from_checkpoint is not None:
        if args.overwrite_output_dir:
            raise SystemExit(
                "--resume-from-checkpoint cannot be combined with --overwrite-output-dir."
            )
        if not args.resume_from_checkpoint.is_dir():
            raise SystemExit(
                f"Checkpoint directory not found: {args.resume_from_checkpoint}"
            )
        required_checkpoint_files = ("trainer_state.json", "optimizer.pt", "scheduler.pt")
        missing_files = [
            name
            for name in required_checkpoint_files
            if not (args.resume_from_checkpoint / name).is_file()
        ]
        if missing_files:
            raise SystemExit(
                f"Invalid or incomplete checkpoint {args.resume_from_checkpoint}; "
                f"missing: {', '.join(missing_files)}"
            )
    if (
        args.output_dir.exists()
        and any(args.output_dir.iterdir())
        and args.resume_from_checkpoint is None
        and not args.overwrite_output_dir
    ):
        raise SystemExit(
            f"Output directory is not empty: {args.output_dir}. "
            "Choose another directory, resume from a checkpoint, or pass "
            "--overwrite-output-dir."
        )

    command = build_command(args)
    print("Launching BGE-M3 dense fine-tuning:\n", subprocess.list2cmdline(command))
    if args.dry_run:
        return

    environment = os.environ.copy()
    environment.setdefault("WANDB_MODE", "disabled")
    if args.cuda_visible_devices:
        environment["CUDA_VISIBLE_DEVICES"] = args.cuda_visible_devices
        print(f"CUDA_VISIBLE_DEVICES={args.cuda_visible_devices}")
    subprocess.run(command, check=True, env=environment)


if __name__ == "__main__":
    main()
