# datademotry2

## JE samples summary workflow

This repository includes a small analysis workflow for `je_samples.xlsx` that produces basic descriptive outputs you can reuse for future GL dumps.

### What it generates

Running the script (or GitHub Actions workflow) produces the following files in `outputs/`:

- `je_samples_summary.txt`: high-level row/column counts and detected date range.
- `column_summary.csv`: per-column null counts, unique counts, and dtypes.
- `numeric_summary.csv`: descriptive stats (count, mean, std, min, median, max, sum) for numeric columns.
- `date_summary.csv`: min/max per detected date-like column.

### Run locally

```bash
python -m pip install -r requirements.txt
python scripts/analyze_je_samples.py --input je_samples.xlsx --output-dir outputs
```

### Run in GitHub Actions

The `JE Samples Summary` workflow runs on demand or when the inputs change. The outputs are uploaded as an artifact named `je-samples-summary` that you can download from the workflow run.
