import argparse
import csv
import math
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import xml.etree.ElementTree as ET


EXPECTED_DIGITS = {digit: math.log10(1 + 1 / digit) for digit in range(1, 10)}
NAMESPACE = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


@dataclass
class SheetData:
    headers: list[str]
    rows: list[list[Any]]


def column_index_from_ref(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref)
    if not letters:
        return 0
    result = 0
    for char in letters.group(1):
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def read_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    try:
        shared = zip_file.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(shared)
    strings: list[str] = []
    for si in root.findall(f"{NAMESPACE}si"):
        text_parts = [node.text or "" for node in si.findall(f".//{NAMESPACE}t")]
        strings.append("".join(text_parts))
    return strings


def parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        inline = cell.find(f"{NAMESPACE}is")
        text_node = inline.find(f".//{NAMESPACE}t") if inline is not None else None
        return text_node.text if text_node is not None else None

    value_node = cell.find(f"{NAMESPACE}v")
    if value_node is None or value_node.text is None:
        return None
    raw = value_node.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return None
    if cell_type in (None, "n"):
        try:
            return float(raw)
        except ValueError:
            return raw
    return raw


def load_first_sheet(path: Path) -> SheetData:
    with zipfile.ZipFile(path) as zip_file:
        sheet_names = sorted(
            name
            for name in zip_file.namelist()
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
        )
        if not sheet_names:
            raise ValueError("No worksheet found in the Excel file.")
        shared_strings = read_shared_strings(zip_file)
        sheet_data = zip_file.read(sheet_names[0])
    root = ET.fromstring(sheet_data)

    rows: dict[int, dict[int, Any]] = {}
    max_col = 0
    for row in root.findall(f".//{NAMESPACE}row"):
        row_idx = int(row.attrib.get("r", "0"))
        row_cells: dict[int, Any] = {}
        for cell in row.findall(f"{NAMESPACE}c"):
            ref = cell.attrib.get("r", "")
            col_idx = column_index_from_ref(ref)
            if col_idx == 0:
                continue
            value = parse_cell_value(cell, shared_strings)
            row_cells[col_idx] = value
            max_col = max(max_col, col_idx)
        if row_idx:
            rows[row_idx] = row_cells

    if 1 not in rows:
        raise ValueError("Unable to locate header row.")
    headers = [
        str(rows[1].get(col_idx) or f"Column{col_idx}").strip()
        for col_idx in range(1, max_col + 1)
    ]
    data_rows: list[list[Any]] = []
    for row_idx in sorted(idx for idx in rows if idx != 1):
        row_values = [rows[row_idx].get(col_idx) for col_idx in range(1, max_col + 1)]
        data_rows.append(row_values)
    return SheetData(headers=headers, rows=data_rows)


def parse_numeric(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return None


def numeric_columns(sheet: SheetData) -> dict[str, list[float]]:
    columns: dict[str, list[float]] = {header: [] for header in sheet.headers}
    for row in sheet.rows:
        for header, value in zip(sheet.headers, row):
            numeric_value = parse_numeric(value)
            if numeric_value is not None:
                columns[header].append(numeric_value)
    return {
        header: values
        for header, values in columns.items()
        if values and not is_date_like_column(header, values)
    }


def is_date_like_column(header: str, values: list[float]) -> bool:
    if "date" in header.lower():
        return True
    if not values:
        return False
    date_like = 0
    for value in values:
        if 20000 <= value <= 60000 and abs(value - round(value)) < 0.01:
            date_like += 1
    return date_like / len(values) >= 0.8


def leading_digits(values: Iterable[float]) -> list[int]:
    digits: list[int] = []
    for value in values:
        if value == 0:
            continue
        value = abs(value)
        exponent = math.floor(math.log10(value))
        leading = int(value / (10**exponent))
        if 1 <= leading <= 9:
            digits.append(leading)
    return digits


def benford_for_column(values: list[float], label: str) -> list[dict[str, Any]]:
    digits = leading_digits(values)
    counts = {digit: digits.count(digit) for digit in range(1, 10)}
    total = sum(counts.values())
    rows: list[dict[str, Any]] = []
    for digit in range(1, 10):
        observed = counts[digit] / total if total else 0
        expected = EXPECTED_DIGITS[digit]
        rows.append(
            {
                "column": label,
                "digit": digit,
                "count": counts[digit],
                "observed": observed,
                "expected": expected,
                "deviation": observed - expected,
            }
        )
    return rows


def summarize_benford(columns: dict[str, list[float]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    digit_detail: list[dict[str, Any]] = []
    for label, values in columns.items():
        digit_detail.extend(benford_for_column(values, label))

    summary: list[dict[str, Any]] = []
    for label in columns:
        column_rows = [row for row in digit_detail if row["column"] == label]
        deviations = [abs(row["deviation"]) for row in column_rows]
        total_values = sum(row["count"] for row in column_rows)
        max_dev_row = max(column_rows, key=lambda row: abs(row["deviation"]))
        summary.append(
            {
                "column": label,
                "total_values": total_values,
                "mad": sum(deviations) / len(deviations) if deviations else 0,
                "max_abs_deviation": max(deviations) if deviations else 0,
                "top_deviation_digit": max_dev_row["digit"],
            }
        )
    summary.sort(key=lambda row: row["mad"], reverse=True)
    return summary, digit_detail


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def svg_bar_chart(
    output_path: Path,
    title: str,
    labels: list[str],
    series: list[tuple[str, list[float], str]],
) -> None:
    width = 900
    height = 450
    margin = 60
    chart_width = width - 2 * margin
    chart_height = height - 2 * margin
    max_value = max(max(values) for _, values, _ in series) if series else 1

    group_count = len(labels)
    series_count = len(series)
    group_width = chart_width / max(group_count, 1)
    bar_width = group_width / max(series_count, 1) * 0.7

    lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        f'<rect width="100%" height="100%" fill="white"/>',
        f'<text x="{width / 2}" y="30" text-anchor="middle" font-size="16">{title}</text>',
    ]

    for idx, label in enumerate(labels):
        x = margin + idx * group_width + group_width / 2
        lines.append(
            f'<text x="{x}" y="{height - margin + 20}" text-anchor="middle" font-size="12">{label}</text>'
        )

    for series_idx, (name, values, color) in enumerate(series):
        for idx, value in enumerate(values):
            bar_height = (value / max_value) * chart_height if max_value else 0
            x = (
                margin
                + idx * group_width
                + (group_width - bar_width * series_count) / 2
                + series_idx * bar_width
            )
            y = height - margin - bar_height
            lines.append(
                f'<rect x="{x:.2f}" y="{y:.2f}" width="{bar_width:.2f}" '
                f'height="{bar_height:.2f}" fill="{color}"/>'
            )

    legend_x = width - margin + 10
    legend_y = margin
    for idx, (name, _, color) in enumerate(series):
        y = legend_y + idx * 20
        lines.append(f'<rect x="{legend_x}" y="{y - 10}" width="12" height="12" fill="{color}"/>')
        lines.append(f'<text x="{legend_x + 18}" y="{y}" font-size="12">{name}</text>')

    lines.append("</svg>")
    output_path.write_text("\n".join(lines), encoding="utf-8")


def plot_overall(digit_detail: list[dict[str, Any]], output_dir: Path) -> None:
    totals = {digit: 0 for digit in range(1, 10)}
    for row in digit_detail:
        totals[row["digit"]] += row["count"]
    total_count = sum(totals.values())
    observed = [(totals[digit] / total_count) if total_count else 0 for digit in range(1, 10)]
    expected = [EXPECTED_DIGITS[digit] for digit in range(1, 10)]
    labels = [str(digit) for digit in range(1, 10)]
    svg_bar_chart(
        output_dir / "benford_overall.svg",
        "Overall Benford Analysis (Observed vs Expected)",
        labels,
        [("Observed", observed, "#1b9e77"), ("Expected", expected, "#7570b3")],
    )


def plot_mad(summary: list[dict[str, Any]], output_dir: Path) -> None:
    labels = [row["column"] for row in summary]
    values = [row["mad"] for row in summary]
    svg_bar_chart(
        output_dir / "benford_mad_by_column.svg",
        "Benford MAD by Column (Higher = More Deviation)",
        labels,
        [("MAD", values, "#d95f02")],
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Benford's analysis on numeric columns.")
    parser.add_argument("--input", default="je_samples.xlsx", help="Path to the input Excel file.")
    parser.add_argument("--output-dir", default="outputs", help="Directory to write output files.")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    sheet = load_first_sheet(Path(args.input))
    columns = numeric_columns(sheet)
    summary, digit_detail = summarize_benford(columns)

    write_csv(output_dir / "benford_summary.csv", summary)
    write_csv(output_dir / "benford_digit_detail.csv", digit_detail)

    if digit_detail:
        plot_overall(digit_detail, output_dir)
    if summary:
        plot_mad(summary, output_dir)


if __name__ == "__main__":
    main()
