# Fine-tuning BGE-M3 for Persian QA Retrieval

این پروژه `BAAI/bge-m3` را به‌صورت dense bi-encoder با Sentence Transformers
فاین‌تیون می‌کند. تمام فرمان‌های پروژه بعد از فعال‌شدن محیط Python در ویندوز و لینوکس
یکسان هستند و آموزش روی یک GPU (`cuda:0`) انجام می‌شود.

## ۱. فعال‌کردن محیط

فقط فعال‌کردن محیط به shell بستگی دارد.

ویندوز PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Linux:

```bash
source .venv/bin/activate
```

یا با Conda:

```bash
conda activate YOUR_ENV
```

از این مرحله به بعد همه فرمان‌ها در هر دو سیستم دقیقاً یکسان‌اند. صحت محیط را بررسی کنید:

```bash
python -c "import torch; print(torch.__version__); print(torch.cuda.is_available()); print(torch.cuda.get_device_name(0))"
```

خروجی CUDA باید `True` باشد. در صورت نیاز:

```bash
python -m pip install -r requirements.txt
```

## ۲. آماده‌سازی داده

فایل `porseman_clean.csv` باید در ریشه پروژه باشد:

```bash
python scripts/prepare_data.py --input porseman_clean.csv --output-dir data/processed --seed 42
python scripts/export_duplicates.py
```

خروجی دیتاست فعلی:

- train: تعداد ۳۷٬۶۷۹
- validation: تعداد ۴٬۷۱۰
- test: تعداد ۴٬۷۱۰
- سؤال تکراری اضافی در داده خام: ۱٬۶۲۴
- جفت کاملاً تکراری در داده خام: ۲۵۸

آمار و hash داده در `data/processed/manifest.json` ذخیره می‌شود.

## ۳. بررسی مدل

```bash
python scripts/inspect_model.py --model BAAI/bge-m3 --output reports/model_inspection.json
```

این فرمان معماری XLM-RoBERTa، ۲۴ لایه، hidden size برابر ۱۰۲۴، تعداد ۱۶ attention
head، CLS pooling، Normalize و حداکثر طول ۸۱۹۲ را کنترل می‌کند.

## ۴. محاسبه MRR و Recall مدل پایه روی test

```bash
python scripts/evaluate.py --model BAAI/bge-m3 --split data/processed/test.jsonl --output reports/baseline_test.json --batch-size 1
```

نتیجه شامل `MRR@10`، `Recall@1/5/10` و `nDCG@10` است و نوار پیشرفت embedding
در ترمینال نمایش داده می‌شود.

برای گزارش ترکیبی معیارها و شباهت آماری train/test:

```bash
python scripts/baseline_report.py --output reports/baseline_report.json --batch-size 1
```

نسخه خوانا در `reports/baseline_report.md` ذخیره می‌شود.

## ۵. smoke test آموزش تک‌GPU

```bash
python scripts/train.py --smoke-test --outer-batch-size 2 --mini-batch-size 1
```

این مرحله دو optimizer step اجرا و checkpoint را مجدداً بارگذاری می‌کند. خروجی در
`outputs/bge-m3-dense/smoke/` ذخیره می‌شود.

warning زیر در مرحله بدون gradient مربوط به cache loss طبیعی است:

```text
None of the inputs have requires_grad=True. Gradients will be None
```

## ۶. آموزش کامل تک‌GPU

فرمان آموزش روی ویندوز و لینوکس یکسان است:

```bash
python scripts/train.py --epochs 1 --outer-batch-size 2 --mini-batch-size 1
```

تنظیمات اصلی:

- یک GPU، به‌صورت پیش‌فرض `cuda:0`
- full fine-tuning و طول ۸۱۹۲ توکن
- FP16 و gradient checkpointing
- `CachedMultipleNegativesRankingLoss`
- Adafactor و learning rate برابر `1e-5`
- انتخاب checkpoint براساس `MRR@10` validation

مدل نهایی در `outputs/bge-m3-dense/run/final` ذخیره می‌شود.

برای ادامه از آخرین checkpoint:

```bash
python scripts/train.py --epochs 1 --outer-batch-size 2 --mini-batch-size 1 --resume-from-checkpoint latest
```

## ۷. ارزیابی مدل فاین‌تیون‌شده

```bash
python scripts/evaluate.py --model outputs/bge-m3-dense/run/final --split data/processed/test.jsonl --output reports/final_test.json --batch-size 1
```

## اجرای یک‌جای pipeline

runner پایتونی مستقل از سیستم‌عامل:

```bash
python scripts/run_pipeline.py
```

اگر داده و baseline قبلاً آماده شده‌اند:

```bash
python scripts/run_pipeline.py --skip-prepare --skip-baseline
```

این runner مراحل را روی یک GPU اجرا و در اولین خطا متوقف می‌کند.

## تست کد

```bash
python -m pytest -q
```
