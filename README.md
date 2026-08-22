# Data Preparation

## Step 0 - Generate Hard Negatives

Input: `data/raw/porseman_clean.csv` with `question` and `content_text` columns

### Current method: separate requests

Generates each of the seven strategies with a separate request. Thinking is enabled by default.

```bash
python scripts/generate_hard_negatives_separate.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_all_separate.jsonl
```

Without thinking:

```bash
python scripts/generate_hard_negatives_separate.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_all_separate.jsonl --no-thinking
```

Custom length tolerance (default: `0.15`):

```bash
python scripts/generate_hard_negatives_separate.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_all_separate.jsonl --length-tolerance 0.10
```

Test a limited number of rows:

```bash
python scripts/generate_hard_negatives_separate.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_test_separate.jsonl --limit 20
```

To resume after interruption, run the same command again.

### Previous method: one combined request

```bash
python scripts/generate_hard_negatives.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_all.jsonl --no-thinking
```

With thinking:

```bash
python scripts/generate_hard_negatives.py --input data/raw/porseman_clean.csv --output data/raw/hard_negatives_all.jsonl --thinking
```

<br>

---

<br>

## Step 1 - Clean and Deduplicate the Dataset

Input: `data/raw/hard_negatives_all_separate.jsonl`

Output: `data/processed/hard_negatives_clean_separate.jsonl`

```bash
python scripts/clean_dataset.py INPUT OUTPUT
```

Example:

```bash
python scripts/clean_dataset.py data/raw/hard_negatives_all_separate.jsonl data/processed/hard_negatives_clean_separate.jsonl
```

This step calls `clean_and_deduplicate_dataset()` to remove duplicate queries,
duplicate negatives, and positives mislabeled as negatives.

<br>

---

<br>

## Step 2 - Create Train, Validation, and Test Files

```bash
python scripts/split_bge_data.py INPUT TRAIN_OUTPUT VALIDATION_OUTPUT TEST_OUTPUT --train-ratio 80 --validation-ratio 10 --test-ratio 10 --seed 42
```

Example:

```bash
python scripts/split_bge_data.py data/processed/hard_negatives_clean_separate.jsonl data/splits/train_separate.jsonl data/splits/validation_separate.jsonl data/splits/test_separate.jsonl --train-ratio 80 --validation-ratio 10 --test-ratio 10 --seed 42
```

Outputs:

- `data/splits/train_separate.jsonl`
- `data/splits/validation_separate.jsonl`
- `data/splits/test_separate.jsonl`

The three ratios must add up to `100`.

<br>

---

<br>

## Step 3 - Evaluate the Test Data

Install dependencies:

```bash
pip install -r requirements.txt
```

Positive-only corpus:

```bash
python scripts/evaluate_positive_only.py TEST_FILE MODEL
```

```bash
python scripts/evaluate_positive_only.py data/splits/test_separate.jsonl BAAI/bge-m3 --devices cuda:0 --fp16
```

Positive and negative corpus:

```bash
python scripts/evaluate_full_corpus.py TEST_FILE MODEL
```

```bash
python scripts/evaluate_full_corpus.py data/splits/test_separate.jsonl BAAI/bge-m3 --devices cuda:0 --fp16
```

Both evaluations report `Recall@1`, `Recall@5`, and `MRR@10`.
Results are saved automatically under a timestamped directory in
`reports/evaluations/`. Use `--output-dir PATH` to choose a custom directory.
Each run creates `report.html`, `report.txt`, and `report.json`.

<br>

---

<br>

## Step 4 - Fine-tune BGE-M3

Input: `data/splits/train_separate.jsonl`

Output: `outputs/bge-m3-dense`

```bash
python scripts/train_bge_m3.py --train-file data/splits/train_separate.jsonl --output-dir outputs/bge-m3-dense --num-gpus 1
```

The script trains the dense BGE-M3 retriever with one positive and seven negatives per query.

Resume training from a checkpoint:

```bash
python scripts/train_bge_m3.py --train-file data/splits/train_separate.jsonl --output-dir outputs/bge-m3-dense --num-gpus 2 --cuda-visible-devices 0,1 --batch-size 1 --passage-max-length 4096 --precision fp16 --resume-from-checkpoint outputs/bge-m3-dense/checkpoint-2000
```

Use the same training arguments as the original run. Do not use `--overwrite-output-dir` when resuming.
The launcher applies the required FlagEmbedding/Transformers resume compatibility fix automatically.

### RTX 3090 (24 GB) Recommended Start

```bash
python scripts/train_bge_m3.py --train-file data/splits/train_separate.jsonl --output-dir outputs/bge-m3-dense --num-gpus 1 --cuda-visible-devices 0 --batch-size 1 --passage-max-length 2048 --precision fp16
```

`--batch-size` is the number of query groups per GPU. Each group contains one positive and seven negatives, so use `1` on an RTX 3090. If CUDA runs out of memory, keep `--batch-size 1` and reduce the passage length first:

```bash
--passage-max-length 1024
```

Key parameters:

- `--cuda-visible-devices 0` selects GPU 0. Use `0,1` with `--num-gpus 2` for two GPUs.
- `--batch-size` is per GPU; default: `1`.
- `--passage-max-length` is the maximum passage token count; default: `2048`.
- `--query-max-length` is the maximum query token count; default: `512`.
- `--epochs`, `--learning-rate`, and `--precision` control the training duration, optimizer rate, and numeric precision.
- `--train-group-size` is one positive plus negatives; default: `8`.
- `--dry-run` prints the underlying FlagEmbedding command without starting training.

Training shows a progress bar and logs metrics every 10 steps. Multi-GPU runs use DDP-safe non-reentrant gradient checkpointing automatically.
