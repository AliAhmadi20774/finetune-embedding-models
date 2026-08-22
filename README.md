# Data Preparation

## Step 1 - Clean and Deduplicate the Dataset

Input: `data/hard_negatives_all.jsonl`

Output: `data/hard_negatives_deduplicated.jsonl`

```bash
python deduplicate_queries.py data/hard_negatives_all.jsonl data/hard_negatives_deduplicated.jsonl
```

This step calls `clean_and_deduplicate_dataset()` to remove duplicate queries,
duplicate negatives, and positives mislabeled as negatives.

<br>

---

<br>

## Step 2 - Create Train, Validation, and Test Files

```bash
python split_bge_data.py INPUT TRAIN_OUTPUT VALIDATION_OUTPUT TEST_OUTPUT --train-ratio 80 --validation-ratio 10 --test-ratio 10 --seed 42
```

Example:

```bash
python split_bge_data.py data/hard_negatives_deduplicated.jsonl data/bge_m3/train.jsonl data/bge_m3/validation.jsonl data/bge_m3/test.jsonl --train-ratio 80 --validation-ratio 10 --test-ratio 10 --seed 42
```

Outputs:

- `data/bge_m3/train.jsonl`
- `data/bge_m3/validation.jsonl`
- `data/bge_m3/test.jsonl`

The three ratios must add up to `100`.
