#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import sqlite3
from pathlib import Path


TYPE_MAP = {
    "2": "INTEGER",
    "3": "INTEGER",
    "4": "SINGLE",
    "5": "REAL",
    "7": "TEXT",
    "11": "INTEGER",
    "130": "TEXT",
}

DEFAULT_INPUT_TABLES = [
    "Cultivars",
    "CO2Yearly",
    "Dweather",
    "FertiMin_List",
    "FertiOrga_List",
    "General_Parameters",
    "Irrigation_List",
    "ListPAnnexes",
    "ListResidus",
    "Mulch",
    "OptionsModel",
    "ParamIni",
    "PlantSpecies",
    "RuissellementObs",
    "SimUnitList",
    "Soil",
    "Soil_layers",
    "StadePheno",
    "Tech_Commun",
    "Tech_perCrop",
    "TypeSurfSol",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build a SQLite input database from TSV exports of Access tables.")
    parser.add_argument("schema_csv", type=Path, help="CSV describing the Access table schema.")
    parser.add_argument("tsv_dir", type=Path, help="Directory containing one TSV file per table.")
    parser.add_argument("sqlite_db", type=Path, help="Output SQLite database path.")
    parser.add_argument(
        "--tables",
        nargs="+",
        default=DEFAULT_INPUT_TABLES,
        help="Table names to import. Defaults to the CELSIUS input tables.",
    )
    parser.add_argument(
        "--empty-tables",
        nargs="+",
        default=["OutputSynt", "OutputD_1", "OutputD_2"],
        help="Tables created from the Access schema without importing rows.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    schema = load_schema(args.schema_csv)

    if args.sqlite_db.exists():
        args.sqlite_db.unlink()

    args.sqlite_db.parent.mkdir(parents=True, exist_ok=True)

    connection = sqlite3.connect(args.sqlite_db)
    try:
        for table in args.tables:
            if table not in schema:
                raise KeyError(f"Table '{table}' was not found in schema file {args.schema_csv}.")
            create_table(connection, table, schema[table])
            load_table(connection, table, args.tsv_dir / f"{table}.tsv", schema[table])
        for table in args.empty_tables:
            if table not in schema:
                raise KeyError(f"Table '{table}' was not found in schema file {args.schema_csv}.")
            create_table(connection, table, schema[table])
        connection.commit()
    finally:
        connection.close()


def load_schema(path: Path) -> dict[str, list[dict[str, str]]]:
    tables: dict[str, list[dict[str, str]]] = {}
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            tables.setdefault(row["TableName"], []).append(row)

    for columns in tables.values():
        columns.sort(key=lambda item: int(item["Ordinal"]))

    return tables


def create_table(connection: sqlite3.Connection, table: str, columns: list[dict[str, str]]) -> None:
    column_sql = []
    for column in columns:
        sqlite_type = TYPE_MAP.get(column["DataType"], "TEXT")
        nullable = "" if column["IsNullable"] == "True" else " NOT NULL"
        column_sql.append(f'[{column["ColumnName"]}] {sqlite_type}{nullable}')

    connection.execute(f'DROP TABLE IF EXISTS [{table}]')
    connection.execute(f'CREATE TABLE [{table}] ({", ".join(column_sql)})')


def load_table(connection: sqlite3.Connection, table: str, tsv_path: Path, columns: list[dict[str, str]]) -> None:
    if not tsv_path.exists():
        raise FileNotFoundError(f"Missing TSV for table '{table}': {tsv_path}")

    names = [column["ColumnName"] for column in columns]
    placeholders = ", ".join(["?"] * len(names))
    insert_sql = f'INSERT INTO [{table}] ({", ".join(f"[{name}]" for name in names)}) VALUES ({placeholders})'

    with tsv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        rows = []
        for row in reader:
            rows.append([normalize(row.get(name, ""), column["DataType"]) for name, column in zip(names, columns)])

    if rows:
        connection.executemany(insert_sql, rows)


def normalize(value: str | None, data_type: str):
    if value is None or value == "":
        return None
    if data_type == "11":
        return 1 if value.lower() in {"true", "yes", "1", "-1"} else 0
    if data_type in {"2", "3"}:
        return int(float(value))
    if data_type in {"4", "5"}:
        return float(value)
    return value


if __name__ == "__main__":
    main()
