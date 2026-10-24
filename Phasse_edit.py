"""Preview or safely change phase IDs for project-scoped database rows.

For PyCharm's Run button, edit the RUN_* constants below. The default is a
dry run: no row is updated unless RUN_EXECUTE is explicitly set to True.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Iterable, Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402

PROJECT_TABLE = "project_info"
PROJECT_KEY = "id"
PHASE_TABLE = "phase"
PHASE_KEY = "id"

# PyCharm Run-button configuration. Always preview first.
RUN_PROJECT_IDS = [1019]
RUN_NEW_PHASE_ID = 33
# None means every current phase belonging to the selected projects.
# RUN_SOURCE_PHASE_ID: int | None = None
RUN_SOURCE_PHASE_ID= 32
RUN_EXECUTE = False
RUN_ASSUME_YES = False

# Reviewed legacy project columns that may not have a direct project_info FK.
# Never infer safety merely from a column being named project_id.
LEGACY_PROJECT_COLUMNS: dict[str, tuple[str, ...]] = {
    "column_resized_tbl": ("project_info_id",),
    "detail_data": ("project_id",),
    "detail_data_design": ("project_id",),
    "detail_data_rfl": ("project_id",),
    "detail_data_rlist": ("project_id",),
    "detail_data_to": ("project_id",),
    "input_surrogate_model_parameter": ("project_id",),
    "map_folder_url": ("project_id",),
    "operational_state": ("project_id",),
    "parameter_map_relation": ("project_id",),
    "prj_rfl": ("project_info_id",),
    "project_parameter": ("project_id",),
    "project_r_parameter": ("project_id",),
    "project_statement": ("project_id",),
    "project_summary": ("project_id",),
    "rfl_order": ("project_id",),
    "rfl_to_description": ("project_info_id",),
    "rfl_to_statement": ("project_info_id",),
    "sampling_data_path": ("project_id",),
    "study_group": ("project_id",),
    "surrogate_model_path": ("project_id",),
    "temp_data": ("project_id",),
    "test_project_senario_parameter": ("project_id",),
    "test_senario_parameter_map_relation": ("project_id",),
    "wp_statement": ("project_id",),
}

IDENTIFIER_PATTERN = re.compile(r"^[A-Za-z0-9_$]+$")


def quote_identifier(identifier: str) -> str:
    """Validate and quote an identifier obtained from trusted metadata."""
    if not IDENTIFIER_PATTERN.fullmatch(identifier):
        raise ValueError(f"Invalid table/column name: {identifier!r}")
    return f"`{identifier}`"


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create SQL placeholders for an empty list.")
    return ", ".join(["%s"] * len(values))


def normalize_positive_ids(values: Iterable[int], *, label: str) -> list[int]:
    """Return positive integer IDs, de-duplicated while preserving order."""
    normalized: list[int] = []
    seen: set[int] = set()
    for raw_value in values:
        if isinstance(raw_value, bool):
            raise ValueError(f"{label} must be a positive integer, not boolean.")
        try:
            value = int(raw_value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid {label}: {raw_value!r}") from exc
        if value <= 0:
            raise ValueError(f"{label} must be greater than zero: {value}")
        if value not in seen:
            normalized.append(value)
            seen.add(value)
    if not normalized:
        raise ValueError(f"The {label} list is empty.")
    return normalized


def normalize_positive_id(value: int, *, label: str) -> int:
    return normalize_positive_ids([value], label=label)[0]


def get_existing_legacy_columns(cursor) -> dict[str, list[str]]:
    """Return only allowlisted table/column pairs present in this database."""
    table_names = sorted(LEGACY_PROJECT_COLUMNS)
    placeholders = make_placeholders(table_names)
    cursor.execute(
        f"""
        SELECT c.TABLE_NAME, c.COLUMN_NAME
        FROM information_schema.COLUMNS AS c
        JOIN information_schema.COLUMNS AS phase_column
          ON phase_column.TABLE_SCHEMA = c.TABLE_SCHEMA
         AND phase_column.TABLE_NAME = c.TABLE_NAME
         AND phase_column.COLUMN_NAME = 'phase_id'
        JOIN information_schema.TABLES AS tables_metadata
          ON tables_metadata.TABLE_SCHEMA = c.TABLE_SCHEMA
         AND tables_metadata.TABLE_NAME = c.TABLE_NAME
         AND tables_metadata.TABLE_TYPE = 'BASE TABLE'
        WHERE c.TABLE_SCHEMA = DATABASE()
          AND c.TABLE_NAME IN ({placeholders})
        """,
        table_names,
    )
    existing: dict[str, list[str]] = defaultdict(list)
    for table_name, column_name in cursor.fetchall():
        if column_name in LEGACY_PROJECT_COLUMNS[table_name]:
            existing[table_name].append(column_name)
    return dict(existing)


def get_phase_tables(cursor) -> list[dict[str, object]]:
    """Discover base tables with direct FKs to both phase and project."""
    cursor.execute(
        """
        SELECT DISTINCT phase_fk.TABLE_NAME, phase_fk.COLUMN_NAME,
                        project_fk.COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE AS phase_fk
        JOIN information_schema.KEY_COLUMN_USAGE AS project_fk
          ON project_fk.TABLE_SCHEMA = phase_fk.TABLE_SCHEMA
         AND project_fk.TABLE_NAME = phase_fk.TABLE_NAME
        JOIN information_schema.TABLES AS tables_metadata
          ON tables_metadata.TABLE_SCHEMA = phase_fk.TABLE_SCHEMA
         AND tables_metadata.TABLE_NAME = phase_fk.TABLE_NAME
        WHERE phase_fk.TABLE_SCHEMA = DATABASE()
          AND tables_metadata.TABLE_TYPE = 'BASE TABLE'
          AND phase_fk.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND phase_fk.REFERENCED_TABLE_NAME = %s
          AND phase_fk.REFERENCED_COLUMN_NAME = %s
          AND project_fk.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND project_fk.REFERENCED_TABLE_NAME = %s
          AND project_fk.REFERENCED_COLUMN_NAME = %s
        ORDER BY phase_fk.TABLE_NAME, project_fk.COLUMN_NAME
        """,
        (PHASE_TABLE, PHASE_KEY, PROJECT_TABLE, PROJECT_KEY),
    )

    phase_columns: dict[str, str] = {}
    project_columns: dict[str, set[str]] = defaultdict(set)
    relation_types: dict[str, set[str]] = defaultdict(set)
    for table_name, phase_column, project_column in cursor.fetchall():
        phase_columns[table_name] = phase_column
        project_columns[table_name].add(project_column)
        relation_types[table_name].add("foreign_key")

    for table_name, columns in get_existing_legacy_columns(cursor).items():
        phase_columns.setdefault(table_name, "phase_id")
        project_columns[table_name].update(columns)
        relation_types[table_name].add("legacy_allowlist")

    return [
        {
            "table_name": table_name,
            "phase_column": phase_columns[table_name],
            "project_columns": sorted(columns),
            "relation_types": sorted(relation_types[table_name]),
        }
        for table_name, columns in sorted(project_columns.items())
    ]


def build_project_filter(
    columns: Sequence[str], project_ids: Sequence[int]
) -> tuple[str, tuple[int, ...]]:
    if not columns:
        raise ValueError("No project column was provided.")
    placeholders = make_placeholders(project_ids)
    clauses = [f"{quote_identifier(column)} IN ({placeholders})" for column in columns]
    return " OR ".join(clauses), tuple(project_ids) * len(columns)


def build_update_filter(
    table: dict[str, object],
    project_ids: Sequence[int],
    source_phase_id: int | None,
    new_phase_id: int,
) -> tuple[str, tuple[int, ...]]:
    project_sql, project_params = build_project_filter(
        list(table["project_columns"]), project_ids
    )
    phase_column = quote_identifier(str(table["phase_column"]))
    if source_phase_id is None:
        phase_sql = f"{phase_column} <> %s"
        phase_params = (new_phase_id,)
    else:
        phase_sql = f"{phase_column} = %s"
        phase_params = (source_phase_id,)
    return f"({project_sql}) AND {phase_sql}", project_params + phase_params


def find_projects(cursor, project_ids: Sequence[int]) -> list[tuple]:
    placeholders = make_placeholders(project_ids)
    cursor.execute(
        f"SELECT {quote_identifier(PROJECT_KEY)}, project_code "
        f"FROM {quote_identifier(PROJECT_TABLE)} "
        f"WHERE {quote_identifier(PROJECT_KEY)} IN ({placeholders}) "
        f"ORDER BY {quote_identifier(PROJECT_KEY)}",
        project_ids,
    )
    return list(cursor.fetchall())


def phase_exists(cursor, phase_id: int) -> bool:
    cursor.execute(
        f"SELECT 1 FROM {quote_identifier(PHASE_TABLE)} "
        f"WHERE {quote_identifier(PHASE_KEY)} = %s",
        (phase_id,),
    )
    return cursor.fetchone() is not None


def count_rows(
    cursor,
    tables: Sequence[dict[str, object]],
    project_ids: Sequence[int],
    source_phase_id: int | None,
    new_phase_id: int,
) -> dict[str, int]:
    counts: dict[str, int] = {}
    for table in tables:
        table_name = str(table["table_name"])
        where_sql, params = build_update_filter(
            table, project_ids, source_phase_id, new_phase_id
        )
        cursor.execute(
            f"SELECT COUNT(*) FROM {quote_identifier(table_name)} "
            f"WHERE {where_sql}",
            params,
        )
        counts[table_name] = int(cursor.fetchone()[0])
    return counts


def print_plan(
    db_connection,
    projects: Sequence[tuple],
    tables: Sequence[dict[str, object]],
    counts: dict[str, int],
    source_phase_id: int | None,
    new_phase_id: int,
) -> None:
    settings = db_connection.settings_dict
    print(f"Database: {settings.get('NAME')}")
    print(f"Host: {settings.get('HOST')}:{settings.get('PORT') or 3306}")
    print("Projects:")
    for project_id, project_code in projects:
        print(f"- {project_id}: {project_code}")
    print(f"Source phase: {source_phase_id if source_phase_id is not None else 'ALL'}")
    print(f"New phase: {new_phase_id}")
    print("Tables selected from schema and reviewed legacy allowlist:")
    for table in tables:
        table_name = str(table["table_name"])
        project_columns = ", ".join(table["project_columns"])
        sources = ", ".join(table["relation_types"])
        print(
            f"- {table_name}: {counts[table_name]} rows; "
            f"project_columns={project_columns}; source={sources}"
        )
    print(f"Total rows to update: {sum(counts.values())}")


def edit_project_phase(
    project_ids: Iterable[int],
    new_phase_id: int,
    *,
    source_phase_id: int | None = None,
    execute: bool = False,
    assume_yes: bool = False,
    db_connection=connection,
) -> dict[str, int]:
    """Preview changes by default; update only with execute and confirmation."""
    normalized_ids = normalize_positive_ids(project_ids, label="project_id")
    new_phase_id = normalize_positive_id(new_phase_id, label="new_phase_id")
    if source_phase_id is not None:
        source_phase_id = normalize_positive_id(
            source_phase_id, label="source_phase_id"
        )
        if source_phase_id == new_phase_id:
            raise ValueError("source_phase_id and new_phase_id must be different.")

    with db_connection.cursor() as cursor:
        projects = find_projects(cursor, normalized_ids)
        found_ids = {int(project[0]) for project in projects}
        missing_ids = [item for item in normalized_ids if item not in found_ids]
        if missing_ids:
            raise ValueError(f"Project IDs not found: {missing_ids}")
        if not phase_exists(cursor, new_phase_id):
            raise ValueError(f"Target phase does not exist: {new_phase_id}")
        if source_phase_id is not None and not phase_exists(cursor, source_phase_id):
            raise ValueError(f"Source phase does not exist: {source_phase_id}")

        tables = get_phase_tables(cursor)
        counts = count_rows(
            cursor, tables, normalized_ids, source_phase_id, new_phase_id
        )
        print_plan(
            db_connection,
            projects,
            tables,
            counts,
            source_phase_id,
            new_phase_id,
        )

    if not execute:
        print("\nDRY RUN: no data was updated. Use --execute to apply the plan.")
        return counts

    database_name = db_connection.settings_dict.get("NAME") or "<unknown>"
    source_label = str(source_phase_id) if source_phase_id is not None else "ALL"
    confirmation = (
        f"UPDATE_PHASE {database_name}:"
        f"{','.join(map(str, normalized_ids))}:{source_label}->{new_phase_id}"
    )
    if not assume_yes:
        answer = input(f"\nType exactly '{confirmation}' to continue: ").strip()
        if answer != confirmation:
            print("Cancelled. No data was updated.")
            return counts

    updated_counts: dict[str, int] = {}
    with transaction.atomic(using=db_connection.alias):
        with db_connection.cursor() as cursor:
            tables = get_phase_tables(cursor)
            for table in tables:
                table_name = str(table["table_name"])
                phase_column = quote_identifier(str(table["phase_column"]))
                where_sql, params = build_update_filter(
                    table, normalized_ids, source_phase_id, new_phase_id
                )
                cursor.execute(
                    f"SELECT COUNT(*) FROM {quote_identifier(table_name)} "
                    f"WHERE {where_sql} FOR UPDATE",
                    params,
                )
                expected = int(cursor.fetchone()[0])
                cursor.execute(
                    f"UPDATE {quote_identifier(table_name)} "
                    f"SET {phase_column} = %s WHERE {where_sql}",
                    (new_phase_id, *params),
                )
                actual = int(cursor.rowcount)
                if actual != expected:
                    raise RuntimeError(
                        f"Unexpected row count in {table_name}; transaction will "
                        f"roll back: expected={expected}, actual={actual}"
                    )
                updated_counts[table_name] = actual
                print(f"{table_name}: updated {actual} rows")

            remaining = count_rows(
                cursor, tables, normalized_ids, source_phase_id, new_phase_id
            )
            not_updated = {name: count for name, count in remaining.items() if count}
            if not_updated:
                raise RuntimeError(
                    "Rows with the old phase remain; transaction will roll back: "
                    f"{not_updated}"
                )

    print("\nPhase update completed successfully.")
    return updated_counts


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Preview or update phase IDs for selected projects."
    )
    parser.add_argument(
        "--project-ids", nargs="+", type=int, help="Overrides RUN_PROJECT_IDS."
    )
    parser.add_argument("--new-phase-id", type=int, help="Overrides RUN_NEW_PHASE_ID.")
    parser.add_argument(
        "--source-phase-id",
        type=int,
        help="Only update this phase; omit to update every non-target phase.",
    )
    parser.add_argument(
        "--execute", action="store_true", help="Apply UPDATE; default is dry-run."
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip confirmation; only for controlled, audited automation.",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    project_ids = args.project_ids or RUN_PROJECT_IDS
    new_phase_id = (
        args.new_phase_id if args.new_phase_id is not None else RUN_NEW_PHASE_ID
    )
    source_phase_id = (
        args.source_phase_id
        if args.source_phase_id is not None
        else RUN_SOURCE_PHASE_ID
    )
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES
    if assume_yes and not execute:
        raise ValueError("RUN_ASSUME_YES/--yes requires execute mode.")
    edit_project_phase(
        project_ids,
        new_phase_id,
        source_phase_id=source_phase_id,
        execute=execute,
        assume_yes=assume_yes,
    )


if __name__ == "__main__":
    main()
