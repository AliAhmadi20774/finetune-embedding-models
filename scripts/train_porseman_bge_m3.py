"""Fine-tune BAAI/bge-m3 on the Porseman minimal embedding dataset.

This script is intentionally tied to:
    data/processed/porseman_embedding_train_minimal.jsonl

The source dataset uses ``query``, ``positive`` and ``negatives``.  The
FlagEmbedding trainer expects ``query``, ``pos`` and ``neg``, with positives
stored as a list.  This script validates and converts the data before launching
the official dense-only BGE-M3 trainer.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SOURCE_DATASET = (
    PROJECT_ROOT / "data" / "processed" / "porseman_embedding_train_minimal.jsonl"
)
PREPARED_DATASET = (
    PROJECT_ROOT / "data" / "processed" / "porseman_embedding_train_minimal_bge.jsonl"
)
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "outputs" / "bge-m3-porseman-minimal"
MODEL_NAME = "BAAI/bge-m3"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Validate the Porseman minimal dataset and fine-tune the dense "
            "representation of BAAI/bge-m3."
        )
    )
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--epochs", type=float, default=1.0)
    parser.add_argument("--learning-rate", type=float, default=1e-5)
    parser.add_argument(
        "--batch-size",
        type=int,
        default=1,
        help="Per-device query-group batch size.",
    )
    parser.add_argument(
        "--gradient-accumulation-steps",
        type=int,
        default=4,
        help="Accumulate gradients to increase the effective batch size.",
    )
    parser.add_argument(
        "--train-group-size",
        type=int,
        default=8,
        choices=range(2, 9),
        metavar="[2-8]",
        help="One positive plus N-1 negatives. The dataset has 7 negatives per query.",
    )
    parser.add_argument("--query-max-length", type=int, default=1024)
    parser.add_argument("--passage-max-length", type=int, default=1024)
    parser.add_argument(
        "--precision",
        choices=("fp16", "bf16", "fp32"),
        default="fp16",
    )
    parser.add_argument("--warmup-ratio", type=float, default=0.1)
    parser.add_argument("--logging-steps", type=int, default=10)
    parser.add_argument("--save-steps", type=int, default=500)
    parser.add_argument("--save-total-limit", type=int, default=2)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument(
        "--num-gpus",
        type=int,
        default=2,
        help="Number of local GPU training processes.",
    )
    parser.add_argument(
        "--cuda-visible-devices",
        default="0,1",
        help="Comma-separated physical GPU indices exposed to the trainer.",
    )
    parser.add_argument(
        "--rebuild-data",
        action="store_true",
        help="Recreate the FlagEmbedding-compatible JSONL file.",
    )
    parser.add_argument(
        "--prepare-only",
        action="store_true",
        help="Validate and convert the dataset, then stop before training.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate/prepare the data and print the training command only.",
    )
    parser.add_argument(
        "--overwrite-output-dir",
        action="store_true",
        help="Allow Hugging Face Trainer to overwrite an existing output directory.",
    )
    return parser.parse_args()


def validate_args(args: argparse.Namespace) -> None:
    positive_values = {
        "epochs": args.epochs,
        "learning-rate": args.learning_rate,
        "batch-size": args.batch_size,
        "gradient-accumulation-steps": args.gradient_accumulation_steps,
        "query-max-length": args.query_max_length,
        "passage-max-length": args.passage_max_length,
        "save-steps": args.save_steps,
        "save-total-limit": args.save_total_limit,
        "num-gpus": args.num_gpus,
    }
    invalid = [name for name, value in positive_values.items() if value <= 0]
    if invalid:
        raise SystemExit(f"These arguments must be positive: {', '.join(invalid)}")
    if not 0 <= args.warmup_ratio < 1:
        raise SystemExit("--warmup-ratio must be in the [0, 1) interval.")

    selected_devices = [
        device.strip()
        for device in args.cuda_visible_devices.split(",")
        if device.strip()
    ]
    if len(selected_devices) != args.num_gpus:
        raise SystemExit(
            "--num-gpus must match the number of values in "
            "--cuda-visible-devices."
        )


def text_value(value: object) -> str:
    return value.strip() if isinstance(value, str) else ""


def prepare_dataset(force: bool) -> tuple[int, int]:
    if not SOURCE_DATASET.is_file():
        raise SystemExit(f"Source dataset not found: {SOURCE_DATASET}")

    source_is_newer = (
        PREPARED_DATASET.exists()
        and SOURCE_DATASET.stat().st_mtime > PREPARED_DATASET.stat().st_mtime
    )
    should_build = force or not PREPARED_DATASET.exists() or source_is_newer

    if not should_build:
        print(f"Using prepared dataset: {PREPARED_DATASET}")
        return count_jsonl_rows(PREPARED_DATASET), 0

    PREPARED_DATASET.parent.mkdir(parents=True, exist_ok=True)
    valid_rows = 0
    invalid_rows = 0

    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            newline="\n",
            dir=PREPARED_DATASET.parent,
            prefix=f".{PREPARED_DATASET.stem}-",
            suffix=".tmp",
            delete=False,
        ) as target:
            temporary_path = Path(target.name)

            with SOURCE_DATASET.open("r", encoding="utf-8-sig") as source:
                for line_number, line in enumerate(source, start=1):
                    if not line.strip():
                        continue

                    try:
                        row = json.loads(line)
                    except json.JSONDecodeError as exc:
                        raise SystemExit(
                            f"Invalid JSON on source line {line_number}: {exc}"
                        ) from exc

                    if not isinstance(row, dict):
                        invalid_rows += 1
                        continue

                    query = text_value(row.get("query"))
                    positive = text_value(row.get("positive"))
                    raw_negatives = row.get("negatives")

                    if not isinstance(raw_negatives, list):
                        invalid_rows += 1
                        continue

                    negatives: list[str] = []
                    seen: set[str] = {positive}
                    for value in raw_negatives:
                        negative = text_value(value)
                        if negative and negative not in seen:
                            negatives.append(negative)
                            seen.add(negative)

                    if not query or not positive or not negatives:
                        invalid_rows += 1
                        continue

                    converted = {
                        "query": query,
                        "pos": [positive],
                        "neg": negatives,
                    }
                    target.write(json.dumps(converted, ensure_ascii=False) + "\n")
                    valid_rows += 1

        if valid_rows == 0:
            raise SystemExit("The source dataset contains no valid training rows.")

        os.replace(temporary_path, PREPARED_DATASET)
        temporary_path = None
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()

    print(f"Prepared dataset: {PREPARED_DATASET}")
    print(f"Valid rows: {valid_rows:,}")
    print(f"Skipped invalid rows: {invalid_rows:,}")
    return valid_rows, invalid_rows


def count_jsonl_rows(path: Path) -> int:
    with path.open("r", encoding="utf-8") as file:
        return sum(1 for line in file if line.strip())


def check_runtime(precision: str, num_gpus: int) -> None:
    missing = [
        name
        for name in ("torch", "transformers", "FlagEmbedding")
        if importlib.util.find_spec(name) is None
    ]
    if missing:
        packages = " ".join(missing)
        raise SystemExit(
            f"Missing packages: {', '.join(missing)}\n"
            f"Install them with:\n  {sys.executable} -m pip install {packages}"
        )

    import torch

    if not torch.cuda.is_available():
        raise SystemExit(
            "CUDA is not available for the selected --cuda-visible-devices. "
            "Check the GPU indices with nvidia-smi and verify that PyTorch has "
            "CUDA support."
        )
    visible_gpu_count = torch.cuda.device_count()
    if visible_gpu_count != num_gpus:
        raise SystemExit(
            f"Expected {num_gpus} visible GPUs, but PyTorch sees "
            f"{visible_gpu_count}. Check --num-gpus and "
            "--cuda-visible-devices."
        )
    if precision == "bf16" and not torch.cuda.is_bf16_supported():
        raise SystemExit("This GPU/PyTorch combination does not support bf16; use fp16.")

    for device_index in range(visible_gpu_count):
        gpu_name = torch.cuda.get_device_name(device_index)
        memory_gib = (
            torch.cuda.get_device_properties(device_index).total_memory / (1024**3)
        )
        print(f"GPU {device_index}: {gpu_name} ({memory_gib:.1f} GiB)")


def build_training_command(args: argparse.Namespace) -> list[str]:
    output_dir = args.output_dir.resolve()
    cache_root = PROJECT_ROOT / "cache"

    command = [
        sys.executable,
        "-m",
        "torch.distributed.run",
        "--nproc_per_node",
        str(args.num_gpus),
        "-m",
        "FlagEmbedding.finetune.embedder.encoder_only.m3",
        "--model_name_or_path",
        MODEL_NAME,
        "--cache_dir",
        str(cache_root / "models"),
        "--train_data",
        str(PREPARED_DATASET),
        "--cache_path",
        str(cache_root / "flagembedding"),
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
        str(output_dir),
        "--learning_rate",
        str(args.learning_rate),
        "--num_train_epochs",
        str(args.epochs),
        "--per_device_train_batch_size",
        str(args.batch_size),
        "--gradient_accumulation_steps",
        str(args.gradient_accumulation_steps),
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
    if args.overwrite_output_dir:
        command.append("--overwrite_output_dir")
    return command


def main() -> None:
    args = parse_args()
    validate_args(args)

    row_count, invalid_count = prepare_dataset(force=args.rebuild_data)
    if row_count == 0 or invalid_count:
        raise SystemExit("Dataset preparation did not produce a fully valid dataset.")
    if args.prepare_only:
        return

    output_dir = args.output_dir.resolve()
    if output_dir.exists() and any(output_dir.iterdir()) and not args.overwrite_output_dir:
        raise SystemExit(
            f"Output directory is not empty: {output_dir}\n"
            "Choose another --output-dir or pass --overwrite-output-dir."
        )

    command = build_training_command(args)
    print("Training command:")
    print(subprocess.list2cmdline(command))
    if args.dry_run:
        return

    # CUDA reads CUDA_VISIBLE_DEVICES when torch initializes, so set it before
    # the runtime check imports torch. The child processes inherit this value.
    os.environ["CUDA_VISIBLE_DEVICES"] = args.cuda_visible_devices
    check_runtime(args.precision, args.num_gpus)
    output_dir.mkdir(parents=True, exist_ok=True)

    environment = os.environ.copy()
    environment.setdefault("WANDB_MODE", "disabled")
    environment.setdefault("TOKENIZERS_PARALLELISM", "false")

    print(f"Model: {MODEL_NAME}")
    print(f"Training rows: {row_count:,}")
    print(f"Output directory: {output_dir}")
    subprocess.run(command, check=True, cwd=PROJECT_ROOT, env=environment)


if __name__ == "__main__":
    main()
