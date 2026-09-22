#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import math
from pathlib import Path


DEFAULT_OUTPUT_TABLES = ["OutputSynt", "OutputD_1", "OutputD_2"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare TSV table exports with a numeric tolerance.")
    parser.add_argument("left_dir", type=Path, help="Reference TSV directory.")
    parser.add_argument("right_dir", type=Path, help="Candidate TSV directory.")
    parser.add_argument(
        "--tables",
        nargs="+",
        default=DEFAULT_OUTPUT_TABLES,
        help="Tables to compare. Defaults to the CELSIUS output tables.",
    )
    parser.add_argument(
        "--float-tolerance",
        type=float,
        default=1e-6,
        help="Absolute tolerance used to normalize numeric columns before comparison.",
    )
    parser.add_argument(
        "--relative-tolerance",
        type=float,
        default=1e-6,
        help="Relative tolerance for numeric values.",
    )
    parser.add_argument(
        "--ignore-column",
        action="append",
        default=[],
        help="Column to ignore as Table.Column; may be repeated.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    failures: list[str] = []

    for table in args.tables:
        left_path = args.left_dir / f"{table}.tsv"
        right_path = args.right_dir / f"{table}.tsv"
        ignored = {
            item.split(".", 1)[1]
            for item in args.ignore_column
            if item.startswith(table + ".") and "." in item
        }
        compare_table(
            table,
            left_path,
            right_path,
            args.float_tolerance,
            args.relative_tolerance,
            ignored,
            failures,
        )

    if failures:
        for failure in failures:
            print(failure)
        raise SystemExit(1)

    print("COMPARE_OK")


def compare_table(
    table: str,
    left_path: Path,
    right_path: Path,
    tolerance: float,
    relative_tolerance: float,
    ignored_columns: set[str],
    failures: list[str],
) -> None:
    left_rows, left_columns = read_tsv(left_path)
    right_rows, right_columns = read_tsv(right_path)

    if left_columns != right_columns:
        failures.append(f"{table}: columns differ.\nLEFT={left_columns}\nRIGHT={right_columns}")
        return

    key_column = left_columns[0]
    left_by_key = {row[key_column]: row for row in left_rows}
    right_by_key = {row[key_column]: row for row in right_rows}
    if len(left_by_key) != len(left_rows) or len(right_by_key) != len(right_rows):
        failures.append(f"{table}: duplicate values in key column {key_column}.")
        return
    if left_by_key.keys() != right_by_key.keys():
        missing = sorted(left_by_key.keys() - right_by_key.keys())
        extra = sorted(right_by_key.keys() - left_by_key.keys())
        failures.append(f"{table}: keys differ. missing={missing[:1]} extra={extra[:1]}")
        return

    numeric_columns = detect_numeric_columns(left_rows + right_rows, left_columns)
    differences: list[tuple[str, str, str, str]] = []
    for key, left_row in left_by_key.items():
        right_row = right_by_key[key]
        for column in left_columns:
            if column in ignored_columns:
                continue
            left_value = (left_row.get(column) or "").strip()
            right_value = (right_row.get(column) or "").strip()
            if values_equal(
                left_value,
                right_value,
                column in numeric_columns,
                tolerance,
                relative_tolerance,
            ):
                continue
            differences.append((key, column, left_value, right_value))
            if len(differences) >= 5:
                break
        if len(differences) >= 5:
            break

    if not differences:
        print(f"{table}: OK ({len(left_rows)} rows)")
        return

    failures.append(f"{table}: row content differs. left={len(left_rows)} rows right={len(right_rows)} rows")
    for key, column, left_value, right_value in differences:
        failures.append(
            f"{table}: key={key!r} column={column!r} reference={left_value!r} candidate={right_value!r}"
        )


def read_tsv(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    if not path.exists():
        raise FileNotFoundError(f"Missing TSV file: {path}")

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        rows = list(reader)
        return rows, list(reader.fieldnames or [])


def detect_numeric_columns(rows: list[dict[str, str]], columns: list[str]) -> set[str]:
    numeric_columns: set[str] = set()
    for column in columns:
        seen_value = False
        numeric = True
        for row in rows:
            value = (row.get(column) or "").strip()
            if value == "":
                continue
            seen_value = True
            try:
                numeric_value(value)
            except ValueError:
                numeric = False
                break
        if seen_value and numeric:
            numeric_columns.add(column)
    return numeric_columns


def values_equal(left: str, right: str, is_numeric: bool, absolute: float, relative: float) -> bool:
    if left == right:
        return True
    if not is_numeric or left == "" or right == "":
        return False
    return math.isclose(
        numeric_value(left),
        numeric_value(right),
        rel_tol=relative,
        abs_tol=absolute,
    )


def numeric_value(value: str) -> float:
    lowered = value.lower()
    if lowered == "true":
        return 1.0
    if lowered == "false":
        return 0.0
    return float(value)


if __name__ == "__main__":
    main()
