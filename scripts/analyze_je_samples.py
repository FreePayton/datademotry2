import argparse
from pathlib import Path

import pandas as pd


def detect_date_columns(frame: pd.DataFrame) -> tuple[list[str], pd.DataFrame]:
    date_columns: list[str] = []
    updated = frame.copy()
    for column in updated.columns:
        series = updated[column]
        if pd.api.types.is_datetime64_any_dtype(series):
            date_columns.append(column)
            continue
        if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
            non_null_count = series.notna().sum()
            if non_null_count == 0:
                continue
            parsed = pd.to_datetime(series, errors="coerce")
            parsed_non_null = parsed.notna().sum()
            if parsed_non_null / non_null_count >= 0.8:
                updated[column] = parsed
                date_columns.append(column)
    return date_columns, updated


def build_summary(frame: pd.DataFrame, output_dir: Path) -> None:
    row_count = len(frame)
    column_count = len(frame.columns)

    date_columns, updated = detect_date_columns(frame)
    numeric_columns = updated.select_dtypes(include="number").columns.tolist()

    column_summary = pd.DataFrame(
        {
            "column": updated.columns,
            "non_null_count": updated.notna().sum().values,
            "null_count": updated.isna().sum().values,
            "unique_count": updated.nunique(dropna=True).values,
            "dtype": updated.dtypes.astype(str).values,
        }
    )
    column_summary.to_csv(output_dir / "column_summary.csv", index=False)

    if numeric_columns:
        numeric_summary = (
            updated[numeric_columns]
            .agg(["count", "mean", "std", "min", "median", "max", "sum"])
            .transpose()
        )
        numeric_summary.to_csv(output_dir / "numeric_summary.csv")
    else:
        numeric_summary = pd.DataFrame()

    if date_columns:
        date_summary = pd.DataFrame(
            {
                "column": date_columns,
                "min": [updated[col].min() for col in date_columns],
                "max": [updated[col].max() for col in date_columns],
            }
        )
        date_summary.to_csv(output_dir / "date_summary.csv", index=False)
        overall_min = date_summary["min"].min()
        overall_max = date_summary["max"].max()
    else:
        date_summary = pd.DataFrame()
        overall_min = None
        overall_max = None

    summary_lines = [
        "JE Samples Summary",
        "===================",
        f"Rows: {row_count}",
        f"Columns: {column_count}",
        "",
        f"Numeric columns ({len(numeric_columns)}): {', '.join(numeric_columns) if numeric_columns else 'None'}",
        f"Date columns ({len(date_columns)}): {', '.join(date_columns) if date_columns else 'None'}",
    ]
    if overall_min is not None and overall_max is not None:
        summary_lines.extend(
            [
                "",
                f"Overall date range: {overall_min} to {overall_max}",
            ]
        )
    summary_lines.extend(
        [
            "",
            "Outputs:",
            "- column_summary.csv (row counts, nulls, uniques, dtypes)",
            "- numeric_summary.csv (descriptive stats for numeric columns)",
            "- date_summary.csv (min/max for date-like columns)",
        ]
    )
    (output_dir / "je_samples_summary.txt").write_text("\n".join(summary_lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze je_samples.xlsx and write summary outputs.")
    parser.add_argument(
        "--input",
        default="je_samples.xlsx",
        help="Path to the input Excel file.",
    )
    parser.add_argument(
        "--output-dir",
        default="outputs",
        help="Directory to write output files.",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    frame = pd.read_excel(input_path)
    build_summary(frame, output_dir)


if __name__ == "__main__":
    main()
