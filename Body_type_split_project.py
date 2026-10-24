"""Preview or safely split one project phase to a different body type.

The tool selects projects by exact positive IDs. When a compatible
setup does not exist, it clones the setup's R/L relations, L-R relations and
RFL rows. It then clones project_info with the target setup and moves every
reviewed phase-scoped row from the source project to the new project. Other
phases remain on the source project and therefore keep their original body
type. Rules may specify an exact body_type_id when names are duplicated. The
default is a dry run.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Iterable, Mapping, Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402

# PyCharm Run-button configuration for splitting one phase to a new project.
RUN_PROJECT_BODY_TYPE_RULES = [
    {
        "project_id": 32,
        "new_project_id": 1021,
        "phase_id": 6,
        "body_type_id": 8,
        "body_type": "CrossoverSUV",
    },
    {
        "project_id": 33,
        "new_project_id": 1022,
        "phase_id": 6,
        "body_type_id": 8,
        "body_type": "CrossoverSUV",
    },
]
RUN_EXECUTE = False
RUN_ASSUME_YES = False

MUTATION_TABLES = (
    "setup",
    "project_info",
    "setup_r_relation",
    "setup_l_relation",
    "l_r_relation",
    "rfl",
    "prj_rfl",
    "detail_data_rfl",
    "detail_data_to",
    "operational_state",
)
INSERT_TABLES = (
    "setup",
    "setup_r_relation",
    "setup_l_relation",
    "l_r_relation",
    "rfl",
)
UPDATE_TABLES = (
    "prj_rfl",
    "detail_data_rfl",
    "detail_data_to",
    "operational_state",
)
ADVISORY_LOCK_TIMEOUT_SECONDS = 10

# Reviewed legacy project columns. Entries are considered only when the table
# also has phase_id; actual tables/columns are verified through information_schema.
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


def normalize_rules(rules: Iterable[Mapping[str, object]]) -> list[dict[str, object]]:
    """Validate project/phase/body-type rules and reject duplicates."""
    normalized: list[dict[str, object]] = []
    seen: dict[tuple[int, int], tuple[int, int | None, str]] = {}
    new_project_ids: set[int] = set()
    for rule in rules:
        raw_project_id = rule.get("project_id")
        raw_new_project_id = rule.get("new_project_id")
        raw_body_type = rule.get("body_type")
        raw_body_type_id = rule.get("body_type_id")
        raw_phase_id = rule.get("phase_id")
        if not isinstance(raw_body_type, str) or not raw_body_type.strip():
            raise ValueError("Body-type names must be non-empty strings.")
        body_type = raw_body_type.strip()
        if isinstance(raw_project_id, bool):
            raise ValueError("project_id must be a positive integer.")
        try:
            project_id = int(raw_project_id)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid project_id: {raw_project_id!r}") from exc
        if project_id <= 0:
            raise ValueError("project_id must be a positive integer.")
        if isinstance(raw_new_project_id, bool):
            raise ValueError("new_project_id must be a positive integer.")
        try:
            new_project_id = int(raw_new_project_id)
        except (TypeError, ValueError) as exc:
            raise ValueError(
                f"Invalid new_project_id: {raw_new_project_id!r}"
            ) from exc
        if new_project_id <= 0:
            raise ValueError("new_project_id must be a positive integer.")
        if new_project_id == project_id:
            raise ValueError("new_project_id must differ from project_id.")
        if isinstance(raw_phase_id, bool):
            raise ValueError("phase_id must be a positive integer.")
        try:
            phase_id = int(raw_phase_id)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid phase_id: {raw_phase_id!r}") from exc
        if phase_id <= 0:
            raise ValueError("phase_id must be a positive integer.")
        if raw_body_type_id in (None, ""):
            body_type_id = None
        else:
            if isinstance(raw_body_type_id, bool):
                raise ValueError("body_type_id must be a positive integer.")
            try:
                body_type_id = int(raw_body_type_id)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"Invalid body_type_id: {raw_body_type_id!r}") from exc
            if body_type_id <= 0:
                raise ValueError("body_type_id must be a positive integer.")
        key = (project_id, phase_id)
        target = (new_project_id, body_type_id, body_type)
        if key in seen and seen[key] != target:
            raise ValueError(
                f"Conflicting body types for project {project_id}, phase {phase_id}."
            )
        if key not in seen:
            if new_project_id in new_project_ids:
                raise ValueError(
                    f"new_project_id is used by more than one rule: {new_project_id}"
                )
            normalized.append(
                {
                    "project_id": project_id,
                    "new_project_id": new_project_id,
                    "phase_id": phase_id,
                    "body_type_id": body_type_id,
                    "body_type": body_type,
                }
            )
            new_project_ids.add(new_project_id)
        seen[key] = target
    if not normalized:
        raise ValueError("The project/body-type mapping is empty.")
    source_project_ids = {int(rule["project_id"]) for rule in normalized}
    overlap = source_project_ids & new_project_ids
    if overlap:
        raise ValueError(
            "new_project_id cannot be another source project ID: "
            f"{sorted(overlap)}"
        )
    return normalized


def parse_rule_items(items: Iterable[str]) -> list[dict[str, object]]:
    """Parse PROJECT:PHASE:NEW_PROJECT[:BODY_TYPE_ID]=BODY_TYPE values."""
    parsed: list[dict[str, object]] = []
    for item in items:
        target, separator, body_type = item.partition("=")
        target_parts = target.split(":")
        if not separator or len(target_parts) not in {3, 4}:
            raise ValueError(
                f"Invalid rule {item!r}; expected "
                "PROJECT_ID:PHASE_ID:NEW_PROJECT_ID[:BODY_TYPE_ID]=BODY_TYPE."
            )
        project_id, phase_id, new_project_id = target_parts[:3]
        parsed_rule: dict[str, object] = {
            "project_id": project_id,
            "phase_id": phase_id,
            "new_project_id": new_project_id,
            "body_type": body_type,
        }
        if len(target_parts) == 4:
            parsed_rule["body_type_id"] = target_parts[3]
        parsed.append(parsed_rule)
    return normalize_rules(parsed)


def rows_as_dicts(cursor) -> list[dict[str, object]]:
    columns = [item[0] for item in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create placeholders for an empty sequence.")
    return ", ".join(["%s"] * len(values))


def quote_identifier(identifier: str) -> str:
    if not IDENTIFIER_PATTERN.fullmatch(identifier):
        raise ValueError(f"Invalid table/column name: {identifier!r}")
    return f"`{identifier}`"


def get_phase_project_tables(cursor) -> list[dict[str, object]]:
    """Discover reviewed base tables containing both project and phase keys."""
    cursor.execute(
        """
        SELECT DISTINCT phase_fk.TABLE_NAME, phase_fk.COLUMN_NAME,
                        project_fk.COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE AS phase_fk
        JOIN information_schema.KEY_COLUMN_USAGE AS project_fk
          ON project_fk.TABLE_SCHEMA = phase_fk.TABLE_SCHEMA
         AND project_fk.TABLE_NAME = phase_fk.TABLE_NAME
        JOIN information_schema.TABLES AS table_metadata
          ON table_metadata.TABLE_SCHEMA = phase_fk.TABLE_SCHEMA
         AND table_metadata.TABLE_NAME = phase_fk.TABLE_NAME
        WHERE phase_fk.TABLE_SCHEMA = DATABASE()
          AND table_metadata.TABLE_TYPE = 'BASE TABLE'
          AND phase_fk.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND phase_fk.REFERENCED_TABLE_NAME = 'phase'
          AND phase_fk.REFERENCED_COLUMN_NAME = 'id'
          AND project_fk.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND project_fk.REFERENCED_TABLE_NAME = 'project_info'
          AND project_fk.REFERENCED_COLUMN_NAME = 'id'
        ORDER BY phase_fk.TABLE_NAME, project_fk.COLUMN_NAME
        """
    )
    phase_columns: dict[str, str] = {}
    project_columns: dict[str, set[str]] = defaultdict(set)
    for table_name, phase_column, project_column in cursor.fetchall():
        phase_columns[str(table_name)] = str(phase_column)
        project_columns[str(table_name)].add(str(project_column))

    legacy_tables = sorted(LEGACY_PROJECT_COLUMNS)
    placeholders = make_placeholders(legacy_tables)
    cursor.execute(
        f"""
        SELECT project_column.TABLE_NAME, project_column.COLUMN_NAME
        FROM information_schema.COLUMNS AS project_column
        JOIN information_schema.COLUMNS AS phase_column
          ON phase_column.TABLE_SCHEMA = project_column.TABLE_SCHEMA
         AND phase_column.TABLE_NAME = project_column.TABLE_NAME
         AND phase_column.COLUMN_NAME = 'phase_id'
        JOIN information_schema.TABLES AS table_metadata
          ON table_metadata.TABLE_SCHEMA = project_column.TABLE_SCHEMA
         AND table_metadata.TABLE_NAME = project_column.TABLE_NAME
         AND table_metadata.TABLE_TYPE = 'BASE TABLE'
        WHERE project_column.TABLE_SCHEMA = DATABASE()
          AND project_column.TABLE_NAME IN ({placeholders})
        """,
        legacy_tables,
    )
    for table_name, project_column in cursor.fetchall():
        table_name = str(table_name)
        project_column = str(project_column)
        if project_column in LEGACY_PROJECT_COLUMNS[table_name]:
            phase_columns.setdefault(table_name, "phase_id")
            project_columns[table_name].add(project_column)

    return [
        {
            "table_name": table_name,
            "phase_column": phase_columns[table_name],
            "project_columns": sorted(columns),
        }
        for table_name, columns in sorted(project_columns.items())
    ]


def count_phase_project_rows(
    cursor,
    tables: Sequence[Mapping[str, object]],
    project_id: int,
    phase_id: int,
) -> dict[tuple[str, str], int]:
    counts: dict[tuple[str, str], int] = {}
    for table in tables:
        table_name = str(table["table_name"])
        phase_column = str(table["phase_column"])
        for project_column in table["project_columns"]:
            project_column = str(project_column)
            cursor.execute(
                f"SELECT COUNT(*) FROM {quote_identifier(table_name)} "
                f"WHERE {quote_identifier(project_column)} = %s "
                f"AND {quote_identifier(phase_column)} = %s",
                (project_id, phase_id),
            )
            counts[(table_name, project_column)] = int(cursor.fetchone()[0])
    return counts


def get_unique_body_type_id(
    cursor, body_type_name: str, requested_body_type_id: int | None = None
) -> int:
    if requested_body_type_id is not None:
        cursor.execute(
            "SELECT body_type FROM body_type WHERE id = %s",
            (requested_body_type_id,),
        )
        rows = cursor.fetchall()
        if len(rows) != 1:
            raise ValueError(f"Body type ID not found: {requested_body_type_id}")
        actual_name = str(rows[0][0]).strip()
        if actual_name != body_type_name:
            raise ValueError(
                f"Body type ID {requested_body_type_id} is {actual_name!r}, "
                f"not {body_type_name!r}."
            )
        return requested_body_type_id

    cursor.execute(
        "SELECT id FROM body_type WHERE body_type = %s ORDER BY id",
        (body_type_name,),
    )
    rows = cursor.fetchall()
    if not rows:
        raise ValueError(f"Body type not found: {body_type_name!r}")
    if len(rows) != 1:
        raise ValueError(
            f"Body type name is ambiguous: {body_type_name!r}; "
            f"matching IDs={[int(row[0]) for row in rows]}"
        )
    return int(rows[0][0])


def find_assignments(
    cursor, rules: Sequence[Mapping[str, object]]
) -> list[dict[str, object]]:
    """Resolve rules to project/phase pairs and reject overlaps."""
    assignments: dict[tuple[int, int], dict[str, object]] = {}
    missing_rules: list[str] = []
    for rule in rules:
        requested_project_id = int(rule["project_id"])
        new_project_id = int(rule["new_project_id"])
        phase_id = int(rule["phase_id"])
        body_type_name = str(rule["body_type"])
        requested_body_type_id = rule.get("body_type_id")
        body_type_id = get_unique_body_type_id(
            cursor,
            body_type_name,
            int(requested_body_type_id) if requested_body_type_id is not None else None,
        )
        cursor.execute("SELECT 1 FROM project_info WHERE id = %s", (new_project_id,))
        if cursor.fetchone() is not None:
            raise ValueError(
                f"new_project_id already exists in project_info: {new_project_id}"
            )
        cursor.execute(
            """
            SELECT p.id AS project_id, p.project_code, p.setup_id AS old_setup_id,
                   s.destination_id, s.body_type_id AS old_body_type_id,
                   old_body.body_type AS old_body_type,
                   s.architecture_id, s.drivetrain_id
            FROM project_info AS p
            JOIN setup AS s ON s.id = p.setup_id
            JOIN body_type AS old_body ON old_body.id = s.body_type_id
            WHERE p.id = %s
              AND (
                    EXISTS (SELECT 1 FROM project_parameter pp
                            WHERE pp.project_id = p.id AND pp.phase_id = %s)
                 OR EXISTS (SELECT 1 FROM project_r_parameter rp
                            WHERE rp.project_id = p.id AND rp.phase_id = %s)
              )
            ORDER BY p.id
            """,
            (requested_project_id, phase_id, phase_id),
        )
        projects = rows_as_dicts(cursor)
        if not projects:
            missing_rules.append(f"project={requested_project_id}:phase={phase_id}")
            continue
        for project in projects:
            project_id = int(project["project_id"])
            cursor.execute(
                """
                SELECT DISTINCT phase_id
                FROM (
                    SELECT phase_id FROM project_parameter WHERE project_id = %s
                    UNION
                    SELECT phase_id FROM project_r_parameter WHERE project_id = %s
                ) AS project_phases
                WHERE phase_id <> %s
                ORDER BY phase_id
                """,
                (project_id, project_id, phase_id),
            )
            other_phase_ids = [int(row[0]) for row in cursor.fetchall()]
            key = (project_id, phase_id)
            previous = assignments.get(key)
            if previous and (
                previous["target_body_type_id"] != body_type_id
                or previous["new_project_id"] != new_project_id
            ):
                raise ValueError(
                    f"Project {project_id}, phase {phase_id} matched rules with "
                    "different target projects or body types."
                )
            project.update(
                {
                    "phase_id": phase_id,
                    "new_project_id": new_project_id,
                    "other_phase_ids": other_phase_ids,
                    "target_body_type": body_type_name,
                    "target_body_type_id": body_type_id,
                }
            )
            assignments[key] = project
    if missing_rules:
        raise ValueError(f"Project/phase data not found for rules: {missing_rules}")
    return [assignments[key] for key in sorted(assignments)]


def find_compatible_setup(cursor, assignment: Mapping[str, object]) -> int | None:
    cursor.execute(
        """
        SELECT id
        FROM setup
        WHERE destination_id = %s
          AND body_type_id = %s
          AND architecture_id = %s
          AND drivetrain_id = %s
        ORDER BY id
        """,
        (
            assignment["destination_id"],
            assignment["target_body_type_id"],
            assignment["architecture_id"],
            assignment["drivetrain_id"],
        ),
    )
    rows = cursor.fetchall()
    if len(rows) > 1:
        raise ValueError(
            "Multiple compatible setups found for project "
            f"{assignment['project_id']}: {[int(row[0]) for row in rows]}"
        )
    return int(rows[0][0]) if rows else None


def count_setup_graph(cursor, setup_id: int) -> dict[str, int]:
    counts: dict[str, int] = {}
    for table in ("setup_r_relation", "setup_l_relation"):
        cursor.execute(f"SELECT COUNT(*) FROM {table} WHERE setup_id = %s", (setup_id,))
        counts[table] = int(cursor.fetchone()[0])
    cursor.execute(
        """
        SELECT COUNT(*)
        FROM l_r_relation AS lr
        JOIN setup_r_relation AS sr ON sr.id = lr.r_s_id
        JOIN setup_l_relation AS sl ON sl.id = lr.l_s_id
        WHERE sr.setup_id = %s AND sl.setup_id = %s
        """,
        (setup_id, setup_id),
    )
    counts["l_r_relation"] = int(cursor.fetchone()[0])
    cursor.execute(
        """
        SELECT COUNT(*)
        FROM rfl
        JOIN setup_r_relation AS sr ON sr.id = rfl.requirement_s_id
        JOIN setup_l_relation AS sl ON sl.id = rfl.logic_s_id
        WHERE sr.setup_id = %s AND sl.setup_id = %s
        """,
        (setup_id, setup_id),
    )
    counts["rfl"] = int(cursor.fetchone()[0])
    return counts


def estimate_insert_counts(
    cursor,
    old_setup_id: int,
    target_setup_id: int | None,
    source_counts: Mapping[str, int],
) -> dict[str, int]:
    """Count setup-graph rows that would need to be inserted."""
    if target_setup_id is None:
        return {
            "setup": 1,
            **{table: int(source_counts[table]) for table in INSERT_TABLES[1:]},
        }

    counts = {"setup": 0}
    for table, parameter_column in (
        ("setup_r_relation", "r_parameter_id"),
        ("setup_l_relation", "l_parameter_id"),
    ):
        cursor.execute(
            f"""
            SELECT COUNT(*)
            FROM {table} AS source_relation
            WHERE source_relation.setup_id = %s
              AND NOT EXISTS (
                    SELECT 1
                    FROM {table} AS target_relation
                    WHERE target_relation.setup_id = %s
                      AND target_relation.{parameter_column}
                          <=> source_relation.{parameter_column}
                      AND target_relation.usecase_id
                          <=> source_relation.usecase_id
              )
            """,
            (old_setup_id, target_setup_id),
        )
        counts[table] = int(cursor.fetchone()[0])

    cursor.execute(
        """
        SELECT COUNT(*)
        FROM l_r_relation AS source_lr
        JOIN setup_r_relation AS source_r ON source_r.id = source_lr.r_s_id
        JOIN setup_l_relation AS source_l ON source_l.id = source_lr.l_s_id
        WHERE source_r.setup_id = %s AND source_l.setup_id = %s
          AND NOT EXISTS (
                SELECT 1
                FROM l_r_relation AS target_lr
                JOIN setup_r_relation AS target_r ON target_r.id = target_lr.r_s_id
                JOIN setup_l_relation AS target_l ON target_l.id = target_lr.l_s_id
                WHERE target_r.setup_id = %s AND target_l.setup_id = %s
                  AND target_r.r_parameter_id <=> source_r.r_parameter_id
                  AND target_r.usecase_id <=> source_r.usecase_id
                  AND target_l.l_parameter_id <=> source_l.l_parameter_id
                  AND target_l.usecase_id <=> source_l.usecase_id
          )
        """,
        (old_setup_id, old_setup_id, target_setup_id, target_setup_id),
    )
    counts["l_r_relation"] = int(cursor.fetchone()[0])

    cursor.execute(
        """
        SELECT COUNT(*)
        FROM rfl AS source_rfl
        JOIN setup_r_relation AS source_r
          ON source_r.id = source_rfl.requirement_s_id
        JOIN setup_l_relation AS source_l ON source_l.id = source_rfl.logic_s_id
        WHERE source_r.setup_id = %s AND source_l.setup_id = %s
          AND NOT EXISTS (
                SELECT 1
                FROM rfl AS target_rfl
                JOIN setup_r_relation AS target_r
                  ON target_r.id = target_rfl.requirement_s_id
                JOIN setup_l_relation AS target_l
                  ON target_l.id = target_rfl.logic_s_id
                WHERE target_r.setup_id = %s AND target_l.setup_id = %s
                  AND target_r.r_parameter_id <=> source_r.r_parameter_id
                  AND target_r.usecase_id <=> source_r.usecase_id
                  AND target_l.l_parameter_id <=> source_l.l_parameter_id
                  AND target_l.usecase_id <=> source_l.usecase_id
                  AND target_rfl.function_id <=> source_rfl.function_id
          )
        """,
        (old_setup_id, old_setup_id, target_setup_id, target_setup_id),
    )
    counts["rfl"] = int(cursor.fetchone()[0])
    return counts


def estimate_update_counts(
    cursor, assignment: Mapping[str, object], target_setup_id: int | None
) -> dict[str, int]:
    """Count phase-scoped rows whose values would actually change."""
    counts: dict[str, int] = {}
    old_setup_id = int(assignment["old_setup_id"])
    if target_setup_id == old_setup_id:
        counts.update({table: 0 for table in UPDATE_TABLES[:-1]})
    else:
        for table, project_column in (
            ("prj_rfl", "project_info_id"),
            ("detail_data_rfl", "project_id"),
            ("detail_data_to", "project_id"),
        ):
            cursor.execute(
                f"""
                SELECT COUNT(*)
                FROM {table} AS reference_row
                JOIN rfl AS source_rfl ON source_rfl.id = reference_row.rfl_id
                JOIN setup_r_relation AS source_r
                  ON source_r.id = source_rfl.requirement_s_id
                JOIN setup_l_relation AS source_l
                  ON source_l.id = source_rfl.logic_s_id
                WHERE reference_row.{project_column} = %s
                  AND reference_row.phase_id = %s
                  AND source_r.setup_id = %s
                  AND source_l.setup_id = %s
                """,
                (
                    assignment["project_id"],
                    assignment["phase_id"],
                    old_setup_id,
                    old_setup_id,
                ),
            )
            counts[table] = int(cursor.fetchone()[0])

    cursor.execute(
        """
        SELECT COUNT(*) FROM operational_state
        WHERE project_id = %s AND phase_id = %s
          AND NOT (destination_id <=> %s AND body_type_id <=> %s
                   AND architecture_id <=> %s AND drivetrain_id <=> %s)
        """,
        (
            assignment["project_id"],
            assignment["phase_id"],
            assignment["destination_id"],
            assignment["target_body_type_id"],
            assignment["architecture_id"],
            assignment["drivetrain_id"],
        ),
    )
    counts["operational_state"] = int(cursor.fetchone()[0])
    return counts


def build_plan(
    cursor, rules: Sequence[Mapping[str, object]]
) -> list[dict[str, object]]:
    plan: list[dict[str, object]] = []
    phase_tables = get_phase_project_tables(cursor)
    for assignment in find_assignments(cursor, rules):
        target_setup_id = find_compatible_setup(cursor, assignment)
        assignment["target_setup_id"] = target_setup_id
        assignment["source_counts"] = count_setup_graph(
            cursor, int(assignment["old_setup_id"])
        )
        assignment["estimated_inserts"] = estimate_insert_counts(
            cursor,
            int(assignment["old_setup_id"]),
            target_setup_id,
            assignment["source_counts"],
        )
        assignment["estimated_updates"] = estimate_update_counts(
            cursor, assignment, target_setup_id
        )
        assignment["phase_tables"] = phase_tables
        assignment["estimated_phase_moves"] = count_phase_project_rows(
            cursor,
            phase_tables,
            int(assignment["project_id"]),
            int(assignment["phase_id"]),
        )
        plan.append(assignment)
    return plan


def print_plan(db_connection, plan: Sequence[Mapping[str, object]]) -> None:
    settings = db_connection.settings_dict
    print(f"Database: {settings.get('NAME')}")
    print(f"Host: {settings.get('HOST')}:{settings.get('PORT') or 3306}")
    phase_table_names = sorted(
        {
            str(table["table_name"])
            for item in plan
            for table in item["phase_tables"]
        }
    )
    affected_tables = list(dict.fromkeys((*MUTATION_TABLES, *phase_table_names)))
    print(f"Affected tables: {', '.join(affected_tables)}")
    print("Projects selected:")
    for item in plan:
        target_setup = item["target_setup_id"]
        setup_action = f"reuse setup {target_setup}" if target_setup else "clone setup"
        print(
            f"- {item['project_id']} {item['project_code']}, phase={item['phase_id']}: "
            f"{item['old_body_type']} -> {item['target_body_type']} "
            f"(body_type_id={item['target_body_type_id']}); "
            f"old_setup={item['old_setup_id']}; {setup_action}; "
            f"new_project_id={item['new_project_id']}"
        )
        print(
            "  New project/Setup/Phase: "
            f"{item['new_project_id']} - "
            f"{target_setup if target_setup is not None else 'NEW_SETUP'} - "
            f"{item['phase_id']}"
        )
        print("  Estimated INSERT rows:")
        print("    project_info: 1")
        for table in INSERT_TABLES:
            print(f"    {table}: {item['estimated_inserts'][table]}")
        if item["other_phase_ids"]:
            print(
                "  Source project keeps other phases: "
                f"{item['other_phase_ids']}"
            )
        else:
            print("  Source project will have no remaining SE/R-list phases.")
        print("  Estimated UPDATE rows:")
        for table in UPDATE_TABLES:
            print(f"    {table}: {item['estimated_updates'][table]}")
        print("  Estimated phase rows moved to the new project:")
        nonzero_moves = [
            (key, count)
            for key, count in item["estimated_phase_moves"].items()
            if count
        ]
        if not nonzero_moves:
            print("    none")
        for (table_name, project_column), count in nonzero_moves:
            print(f"    {table_name}.{project_column}: {count}")


def plan_fingerprint(plan: Sequence[Mapping[str, object]]) -> list[tuple[object, ...]]:
    """Capture fields that must not change between preview and execution."""
    return [
        (
            item["project_id"],
            item["project_code"],
            item["phase_id"],
            item["new_project_id"],
            item["old_setup_id"],
            item["target_setup_id"],
            item["old_body_type_id"],
            item["target_body_type_id"],
            item["destination_id"],
            item["architecture_id"],
            item["drivetrain_id"],
            tuple(item["other_phase_ids"]),
            tuple(sorted(item["estimated_phase_moves"].items())),
        )
        for item in plan
    ]


def acquire_advisory_lock(cursor, database_name: str) -> str:
    lock_name = f"spdm_body_type_edit:{database_name}"[:64]
    cursor.execute(
        "SELECT GET_LOCK(%s, %s)", (lock_name, ADVISORY_LOCK_TIMEOUT_SECONDS)
    )
    result = cursor.fetchone()
    if not result or result[0] != 1:
        raise RuntimeError(f"Could not acquire maintenance lock {lock_name!r}.")
    return lock_name


def release_advisory_lock(cursor, lock_name: str) -> None:
    cursor.execute("SELECT RELEASE_LOCK(%s)", (lock_name,))


def allocate_ids(cursor, table_name: str, count: int) -> list[int]:
    """Allocate IDs while locking the current end of the reviewed table."""
    if table_name not in {
        "setup",
        "setup_r_relation",
        "setup_l_relation",
        "l_r_relation",
        "rfl",
    }:
        raise ValueError(f"ID allocation is not allowed for table {table_name!r}.")
    if count < 0:
        raise ValueError("ID allocation count cannot be negative.")
    if count == 0:
        return []
    cursor.execute(f"SELECT id FROM {table_name} ORDER BY id DESC LIMIT 1 FOR UPDATE")
    row = cursor.fetchone()
    first_id = (int(row[0]) if row and row[0] is not None else 0) + 1
    return list(range(first_id, first_id + count))


def create_setup(cursor, assignment: Mapping[str, object]) -> int:
    new_setup_id = allocate_ids(cursor, "setup", 1)[0]
    cursor.execute(
        """
        INSERT INTO setup
            (id, destination_id, body_type_id, architecture_id, drivetrain_id)
        VALUES (%s, %s, %s, %s, %s)
        """,
        (
            new_setup_id,
            assignment["destination_id"],
            assignment["target_body_type_id"],
            assignment["architecture_id"],
            assignment["drivetrain_id"],
        ),
    )
    if cursor.rowcount != 1:
        raise RuntimeError("Unexpected setup INSERT row count.")
    return new_setup_id


def clone_setup_relations(
    cursor,
    table_name: str,
    parameter_column: str,
    old_setup_id: int,
    new_setup_id: int,
) -> tuple[dict[int, int], int]:
    if (table_name, parameter_column) not in {
        ("setup_r_relation", "r_parameter_id"),
        ("setup_l_relation", "l_parameter_id"),
    }:
        raise ValueError("Unexpected setup relation table/column.")
    cursor.execute(
        f"SELECT id, {parameter_column}, usecase_id FROM {table_name} "
        "WHERE setup_id = %s ORDER BY id",
        (old_setup_id,),
    )
    source_rows = cursor.fetchall()
    result: dict[int, int] = {}
    missing: list[tuple[int, object, object]] = []
    for old_id, parameter_id, usecase_id in source_rows:
        cursor.execute(
            f"SELECT id FROM {table_name} WHERE setup_id = %s "
            f"AND {parameter_column} <=> %s AND usecase_id <=> %s ORDER BY id",
            (new_setup_id, parameter_id, usecase_id),
        )
        matches = cursor.fetchall()
        if len(matches) > 1:
            raise RuntimeError(f"Ambiguous matching rows in {table_name}.")
        if matches:
            result[int(old_id)] = int(matches[0][0])
        else:
            missing.append((int(old_id), parameter_id, usecase_id))
    new_ids = allocate_ids(cursor, table_name, len(missing))
    for new_id, (old_id, parameter_id, usecase_id) in zip(new_ids, missing):
        cursor.execute(
            f"INSERT INTO {table_name} "
            f"(id, setup_id, {parameter_column}, usecase_id) VALUES (%s, %s, %s, %s)",
            (new_id, new_setup_id, parameter_id, usecase_id),
        )
        if cursor.rowcount != 1:
            raise RuntimeError(f"Unexpected INSERT row count in {table_name}.")
        result[old_id] = new_id
    if len(result) != len(source_rows):
        raise RuntimeError(f"Incomplete relation mapping for {table_name}.")
    return result, len(missing)


def clone_l_r_relations(
    cursor, r_map: Mapping[int, int], l_map: Mapping[int, int]
) -> int:
    if not r_map or not l_map:
        return 0
    r_ids = list(r_map)
    l_ids = list(l_map)
    cursor.execute(
        f"SELECT r_s_id, l_s_id FROM l_r_relation "
        f"WHERE r_s_id IN ({make_placeholders(r_ids)}) "
        f"AND l_s_id IN ({make_placeholders(l_ids)}) ORDER BY id",
        (*r_ids, *l_ids),
    )
    target_pairs = [
        (r_map[int(r_id)], l_map[int(l_id)]) for r_id, l_id in cursor.fetchall()
    ]
    missing: list[tuple[int, int]] = []
    for r_id, l_id in target_pairs:
        cursor.execute(
            "SELECT id FROM l_r_relation WHERE r_s_id = %s AND l_s_id = %s",
            (r_id, l_id),
        )
        if cursor.fetchone() is None:
            missing.append((r_id, l_id))
    new_ids = allocate_ids(cursor, "l_r_relation", len(missing))
    for new_id, (r_id, l_id) in zip(new_ids, missing):
        cursor.execute(
            "INSERT INTO l_r_relation (id, r_s_id, l_s_id) VALUES (%s, %s, %s)",
            (new_id, r_id, l_id),
        )
        if cursor.rowcount != 1:
            raise RuntimeError("Unexpected l_r_relation INSERT row count.")
    return len(missing)


def clone_rfl_rows(
    cursor, r_map: Mapping[int, int], l_map: Mapping[int, int]
) -> tuple[dict[int, int], int]:
    if not r_map or not l_map:
        return {}, 0
    r_ids = list(r_map)
    l_ids = list(l_map)
    cursor.execute(
        f"""
        SELECT id, requirement_s_id, function_id, logic_s_id, `index`,
               requirement_condition, function_condition, logic_condition, wp_id
        FROM rfl
        WHERE requirement_s_id IN ({make_placeholders(r_ids)})
          AND logic_s_id IN ({make_placeholders(l_ids)})
        ORDER BY id
        """,
        (*r_ids, *l_ids),
    )
    source_rows = cursor.fetchall()
    result: dict[int, int] = {}
    missing: list[tuple] = []
    for row in source_rows:
        old_id, old_r_id, function_id, old_l_id, *attributes = row
        new_r_id = r_map[int(old_r_id)]
        new_l_id = l_map[int(old_l_id)]
        cursor.execute(
            """
            SELECT id FROM rfl
            WHERE requirement_s_id = %s AND function_id = %s AND logic_s_id = %s
            """,
            (new_r_id, function_id, new_l_id),
        )
        match = cursor.fetchone()
        if match:
            result[int(old_id)] = int(match[0])
        else:
            missing.append((int(old_id), new_r_id, function_id, new_l_id, *attributes))
    new_ids = allocate_ids(cursor, "rfl", len(missing))
    for new_id, row in zip(new_ids, missing):
        old_id, new_r_id, function_id, new_l_id, *attributes = row
        cursor.execute(
            """
            INSERT INTO rfl
                (id, requirement_s_id, function_id, logic_s_id, `index`,
                 requirement_condition, function_condition, logic_condition, wp_id)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (new_id, new_r_id, function_id, new_l_id, *attributes),
        )
        if cursor.rowcount != 1:
            raise RuntimeError("Unexpected rfl INSERT row count.")
        result[old_id] = new_id
    if len(result) != len(source_rows):
        raise RuntimeError("Incomplete RFL mapping.")
    return result, len(missing)


def update_rfl_references(
    cursor, project_id: int, phase_id: int, rfl_map: Mapping[int, int]
) -> dict[str, int]:
    counts: dict[str, int] = {}
    for table_name, project_column in (
        ("prj_rfl", "project_info_id"),
        ("detail_data_rfl", "project_id"),
        ("detail_data_to", "project_id"),
    ):
        updated = 0
        for old_id, new_id in rfl_map.items():
            if old_id == new_id:
                continue
            cursor.execute(
                f"SELECT COUNT(*) FROM {table_name} "
                f"WHERE {project_column} = %s AND phase_id = %s AND rfl_id = %s",
                (project_id, phase_id, old_id),
            )
            expected = int(cursor.fetchone()[0])
            cursor.execute(
                f"UPDATE {table_name} SET rfl_id = %s "
                f"WHERE {project_column} = %s AND phase_id = %s AND rfl_id = %s",
                (new_id, project_id, phase_id, old_id),
            )
            if cursor.rowcount != expected:
                raise RuntimeError(
                    f"Unexpected UPDATE row count in {table_name}: "
                    f"expected={expected}, actual={cursor.rowcount}"
                )
            updated += int(cursor.rowcount)
        counts[table_name] = updated
    return counts


def update_phase_operational_state(
    cursor, assignment: Mapping[str, object]
) -> dict[str, int]:
    project_id = int(assignment["project_id"])
    phase_id = int(assignment["phase_id"])
    dimensions = (
        assignment["destination_id"],
        assignment["target_body_type_id"],
        assignment["architecture_id"],
        assignment["drivetrain_id"],
    )
    cursor.execute(
        """
        SELECT COUNT(*) FROM operational_state
        WHERE project_id = %s AND phase_id = %s
          AND NOT (destination_id <=> %s AND body_type_id <=> %s
                   AND architecture_id <=> %s AND drivetrain_id <=> %s)
        """,
        (project_id, phase_id, *dimensions),
    )
    expected_states = int(cursor.fetchone()[0])
    cursor.execute(
        """
        UPDATE operational_state
        SET destination_id = %s, body_type_id = %s,
            architecture_id = %s, drivetrain_id = %s
        WHERE project_id = %s AND phase_id = %s
          AND NOT (destination_id <=> %s AND body_type_id <=> %s
                   AND architecture_id <=> %s AND drivetrain_id <=> %s)
        """,
        (*dimensions, project_id, phase_id, *dimensions),
    )
    if cursor.rowcount != expected_states:
        raise RuntimeError(
            "Unexpected operational_state UPDATE row count: "
            f"expected={expected_states}, actual={cursor.rowcount}"
        )
    return {
        "operational_state": int(cursor.rowcount),
    }


def create_phase_project(
    cursor, assignment: Mapping[str, object], new_setup_id: int
) -> int:
    """Clone project metadata while replacing only its project and setup IDs."""
    source_project_id = int(assignment["project_id"])
    new_project_id = int(assignment["new_project_id"])
    cursor.execute(
        "SELECT id FROM project_info WHERE id = %s FOR UPDATE",
        (source_project_id,),
    )
    if cursor.fetchone() is None:
        raise RuntimeError(f"Source project disappeared: {source_project_id}")
    cursor.execute("SELECT 1 FROM project_info WHERE id = %s", (new_project_id,))
    if cursor.fetchone() is not None:
        raise RuntimeError(
            f"Target project ID was created after preview: {new_project_id}"
        )
    cursor.execute(
        """
        INSERT INTO project_info
            (id, project_code, setup_id, lot_id, new_or_not, family_id)
        SELECT %s, project_code, %s, lot_id, new_or_not, family_id
        FROM project_info
        WHERE id = %s
        """,
        (new_project_id, new_setup_id, source_project_id),
    )
    if cursor.rowcount != 1:
        raise RuntimeError(
            "Unexpected project_info INSERT row count: "
            f"expected=1, actual={cursor.rowcount}"
        )
    return int(cursor.rowcount)


def move_phase_rows_to_new_project(
    cursor,
    assignment: Mapping[str, object],
) -> dict[tuple[str, str], int]:
    """Move all reviewed rows for the selected phase to the cloned project."""
    source_project_id = int(assignment["project_id"])
    new_project_id = int(assignment["new_project_id"])
    phase_id = int(assignment["phase_id"])
    expected_counts = assignment["estimated_phase_moves"]
    moved: dict[tuple[str, str], int] = {}
    for table in assignment["phase_tables"]:
        table_name = str(table["table_name"])
        phase_column = str(table["phase_column"])
        for project_column_value in table["project_columns"]:
            project_column = str(project_column_value)
            key = (table_name, project_column)
            expected = int(expected_counts[key])
            cursor.execute(
                f"UPDATE {quote_identifier(table_name)} "
                f"SET {quote_identifier(project_column)} = %s "
                f"WHERE {quote_identifier(project_column)} = %s "
                f"AND {quote_identifier(phase_column)} = %s",
                (new_project_id, source_project_id, phase_id),
            )
            if cursor.rowcount != expected:
                raise RuntimeError(
                    f"Unexpected project move count in {table_name}."
                    f"{project_column}: expected={expected}, "
                    f"actual={cursor.rowcount}"
                )
            moved[key] = int(cursor.rowcount)
    return moved


def apply_assignment(cursor, assignment: Mapping[str, object]) -> dict[str, object]:
    old_setup_id = int(assignment["old_setup_id"])
    new_setup_id = find_compatible_setup(cursor, assignment)
    setup_created = 0
    if new_setup_id is None:
        new_setup_id = create_setup(cursor, assignment)
        setup_created = 1
    r_map, r_inserted = clone_setup_relations(
        cursor, "setup_r_relation", "r_parameter_id", old_setup_id, new_setup_id
    )
    l_map, l_inserted = clone_setup_relations(
        cursor, "setup_l_relation", "l_parameter_id", old_setup_id, new_setup_id
    )
    lr_inserted = clone_l_r_relations(cursor, r_map, l_map)
    rfl_map, rfl_inserted = clone_rfl_rows(cursor, r_map, l_map)
    project_inserted = create_phase_project(cursor, assignment, new_setup_id)
    rfl_updates = update_rfl_references(
        cursor, int(assignment["project_id"]), int(assignment["phase_id"]), rfl_map
    )
    final_updates = update_phase_operational_state(cursor, assignment)
    phase_moves = move_phase_rows_to_new_project(cursor, assignment)
    return {
        "target_setup_id": new_setup_id,
        "setup_inserted": setup_created,
        "project_info_inserted": project_inserted,
        "setup_r_relation_inserted": r_inserted,
        "setup_l_relation_inserted": l_inserted,
        "l_r_relation_inserted": lr_inserted,
        "rfl_inserted": rfl_inserted,
        **rfl_updates,
        **final_updates,
        "phase_moves": phase_moves,
    }


def print_actual_counts(
    assignment: Mapping[str, object], counts: Mapping[str, object]
) -> None:
    print(
        f"\nApplied source project {assignment['project_id']} "
        f"{assignment['project_code']}, phase={assignment['phase_id']} "
        f"to new project {assignment['new_project_id']}:"
    )
    print(f"  project_info INSERT rows: {counts['project_info_inserted']}")
    print("  INSERT rows:")
    for table in INSERT_TABLES:
        print(f"    {table}: {counts[f'{table}_inserted']}")
    print("  UPDATE rows:")
    for table in UPDATE_TABLES:
        print(f"    {table}: {counts[table]}")
    print("  Phase rows moved:")
    for (table_name, project_column), count in counts["phase_moves"].items():
        if count:
            print(f"    {table_name}.{project_column}: {count}")


def edit_project_body_types(
    rules: Iterable[Mapping[str, object]],
    *,
    execute: bool = False,
    assume_yes: bool = False,
    db_connection=connection,
) -> list[dict[str, object]]:
    normalized_rules = normalize_rules(rules)
    with db_connection.cursor() as cursor:
        plan = build_plan(cursor, normalized_rules)
        print_plan(db_connection, plan)
    if not execute:
        print("\nDRY RUN: no data was changed. Set RUN_EXECUTE=True to apply.")
        return plan

    database_name = db_connection.settings_dict.get("NAME") or "<unknown>"
    targets = ",".join(
        f"{item['project_id']}->{item['new_project_id']}@{item['phase_id']}"
        for item in plan
    )
    confirmation = f"UPDATE_BODY_TYPE {database_name}:{targets}"
    if not assume_yes:
        answer = input(f"\nType exactly '{confirmation}' to continue: ").strip()
        if answer != confirmation:
            print("Cancelled. No data was changed.")
            return plan

    lock_name: str | None = None
    try:
        with db_connection.cursor() as cursor:
            lock_name = acquire_advisory_lock(cursor, str(database_name))
        with transaction.atomic(using=db_connection.alias):
            with db_connection.cursor() as cursor:
                current_plan = build_plan(cursor, normalized_rules)
                if plan_fingerprint(current_plan) != plan_fingerprint(plan):
                    raise RuntimeError(
                        "Selected project data changed after preview; transaction will "
                        "roll back."
                    )
                applied_counts: dict[int, dict[str, object]] = {}
                for assignment in current_plan:
                    counts = apply_assignment(cursor, assignment)
                    applied_counts[int(assignment["new_project_id"])] = counts
                    print_actual_counts(assignment, counts)
                for assignment in current_plan:
                    actual_setup_id = applied_counts[
                        int(assignment["new_project_id"])
                    ]["target_setup_id"]
                    cursor.execute(
                        """
                        SELECT COUNT(*) FROM operational_state
                        WHERE project_id = %s AND phase_id = %s
                          AND NOT (destination_id <=> %s AND body_type_id <=> %s
                                   AND architecture_id <=> %s
                                   AND drivetrain_id <=> %s)
                        """,
                        (
                            assignment["new_project_id"],
                            assignment["phase_id"],
                            assignment["destination_id"],
                            assignment["target_body_type_id"],
                            assignment["architecture_id"],
                            assignment["drivetrain_id"],
                        ),
                    )
                    if int(cursor.fetchone()[0]) != 0:
                        raise RuntimeError(
                            "Operational-state verification failed for project "
                            f"{assignment['new_project_id']}; transaction will roll back."
                        )
                    cursor.execute(
                        "SELECT COUNT(*) FROM project_info "
                        "WHERE id = %s AND setup_id = %s",
                        (
                            assignment["new_project_id"],
                            actual_setup_id,
                        ),
                    )
                    if int(cursor.fetchone()[0]) != 1:
                        raise RuntimeError(
                            "New project/setup verification failed for project "
                            f"{assignment['new_project_id']}; transaction will roll back."
                        )
                    source_counts = count_phase_project_rows(
                        cursor,
                        assignment["phase_tables"],
                        int(assignment["project_id"]),
                        int(assignment["phase_id"]),
                    )
                    remaining = {
                        key: count for key, count in source_counts.items() if count
                    }
                    if remaining:
                        raise RuntimeError(
                            "Source phase rows remain after project split: "
                            f"{remaining}; transaction will roll back."
                        )
                    target_counts = count_phase_project_rows(
                        cursor,
                        assignment["phase_tables"],
                        int(assignment["new_project_id"]),
                        int(assignment["phase_id"]),
                    )
                    expected_counts = assignment["estimated_phase_moves"]
                    if target_counts != expected_counts:
                        raise RuntimeError(
                            "New project phase-row verification failed; transaction "
                            "will roll back."
                        )
        print("\nBody-type update completed successfully.")
        return plan
    finally:
        if lock_name is not None:
            with db_connection.cursor() as cursor:
                release_advisory_lock(cursor, lock_name)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Preview or safely change body types for project/phase IDs."
    )
    parser.add_argument(
        "--rule",
        action="append",
        metavar="PROJECT_ID:PHASE_ID:NEW_PROJECT_ID[:BODY_TYPE_ID]=BODY_TYPE",
        help=(
            "Repeat for each project/phase split. NEW_PROJECT_ID must not exist. "
            "Supply BODY_TYPE_ID when the body-type name is duplicated; "
            "overrides RUN rules."
        ),
    )
    parser.add_argument(
        "--execute", action="store_true", help="Apply changes; default is dry-run."
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip confirmation; only for controlled, audited automation.",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    rules = (
        parse_rule_items(args.rule)
        if args.rule is not None
        else RUN_PROJECT_BODY_TYPE_RULES
    )
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES
    if assume_yes and not execute:
        raise ValueError("RUN_ASSUME_YES/--yes requires execute mode.")
    edit_project_body_types(rules, execute=execute, assume_yes=assume_yes)


if __name__ == "__main__":
    main()
