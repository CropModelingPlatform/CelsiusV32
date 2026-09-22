#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import sqlite3
from pathlib import Path


DEFAULT_OUTPUT_TABLES = ["OutputSynt", "OutputD_1", "OutputD_2"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export SQLite tables to TSV files.")
    parser.add_argument("sqlite_db", type=Path, help="Path to the SQLite database.")
    parser.add_argument("output_dir", type=Path, help="Directory receiving TSV exports.")
    parser.add_argument(
        "--tables",
        nargs="+",
        default=DEFAULT_OUTPUT_TABLES,
        help="Tables to export. Defaults to the CELSIUS output tables.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    args.output_dir.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(args.sqlite_db)
    try:
        for table in args.tables:
            export_table(connection, table, args.output_dir / f"{table}.tsv")
    finally:
        connection.close()


def export_table(connection: sqlite3.Connection, table: str, output_path: Path) -> None:
    cursor = connection.execute(f"SELECT * FROM [{table}]")
    rows = cursor.fetchall()
    columns = [item[0] for item in cursor.description]

    with output_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t", lineterminator="\n")
        writer.writerow(columns)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])


if __name__ == "__main__":
    main()
