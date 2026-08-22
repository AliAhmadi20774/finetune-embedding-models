# Data Preparation

## Step 1 - Clean and Deduplicate the Dataset

Input: `data/raw/hard_negatives_all.jsonl`

Output: `data/processed/hard_negatives_clean.jsonl`

```bash
python scripts/clean_dataset.py INPUT OUTPUT
```

Example:

```bash
python scripts/clean_dataset.py data/raw/hard_negatives_all.jsonl data/processed/hard_negatives_clean.jsonl
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
python scripts/split_bge_data.py data/processed/hard_negatives_clean.jsonl data/splits/train.jsonl data/splits/validation.jsonl data/splits/test.jsonl --train-ratio 80 --validation-ratio 10 --test-ratio 10 --seed 42
```

Outputs:

- `data/splits/train.jsonl`
- `data/splits/validation.jsonl`
- `data/splits/test.jsonl`

The three ratios must add up to `100`.
