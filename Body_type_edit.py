"""Preview or change setup/RFL data for one project phase and body type.

This is the original in-place workflow: project_info.setup_id is not changed
and no project_info row is created. Only phase-scoped RFL references and
operational_state dimensions are updated. The default is a dry run.
"""

from __future__ import annotations

import argparse
import sys
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

from utils.Tools import Body_type_split_project as core  # noqa: E402

# PyCharm Run-button configuration. Preview before changing RUN_EXECUTE.
RUN_PROJECT_BODY_TYPE_RULES = [
    {
        "project_id": 33,
        "phase_id": 6,
        "body_type_id": 8,
        "body_type": "CrossoverSUV",
    },
    {
        "project_id": 32,
        "phase_id": 6,
        "body_type_id": 8,
        "body_type": "CrossoverSUV",
    },
]
RUN_EXECUTE = False
RUN_ASSUME_YES = False

MUTATION_TABLES = (
    "setup",
    "setup_r_relation",
    "setup_l_relation",
    "l_r_relation",
    "rfl",
    "prj_rfl",
    "detail_data_rfl",
    "detail_data_to",
    "operational_state",
)
INSERT_TABLES = core.INSERT_TABLES
UPDATE_TABLES = core.UPDATE_TABLES


def _positive_id(raw_value: object, label: str) -> int:
    if isinstance(raw_value, bool):
        raise ValueError(f"{label} must be a positive integer.")
    try:
        value = int(raw_value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid {label}: {raw_value!r}") from exc
    if value <= 0:
        raise ValueError(f"{label} must be a positive integer.")
    return value


def normalize_rules(rules: Iterable[Mapping[str, object]]) -> list[dict[str, object]]:
    """Validate project/phase/body-type rules and reject conflicts."""
    normalized: list[dict[str, object]] = []
    seen: dict[tuple[int, int], tuple[int | None, str]] = {}
    for rule in rules:
        project_id = _positive_id(rule.get("project_id"), "project_id")
        phase_id = _positive_id(rule.get("phase_id"), "phase_id")
        raw_body_type = rule.get("body_type")
        if not isinstance(raw_body_type, str) or not raw_body_type.strip():
            raise ValueError("Body-type names must be non-empty strings.")
        body_type = raw_body_type.strip()
        raw_body_type_id = rule.get("body_type_id")
        body_type_id = (
            None
            if raw_body_type_id in (None, "")
            else _positive_id(raw_body_type_id, "body_type_id")
        )
        key = (project_id, phase_id)
        target = (body_type_id, body_type)
        if key in seen and seen[key] != target:
            raise ValueError(
                f"Conflicting body types for project {project_id}, phase {phase_id}."
            )
        if key not in seen:
            normalized.append(
                {
                    "project_id": project_id,
                    "phase_id": phase_id,
                    "body_type_id": body_type_id,
                    "body_type": body_type,
                }
            )
        seen[key] = target
    if not normalized:
        raise ValueError("The project/body-type mapping is empty.")
    return normalized


def parse_rule_items(items: Iterable[str]) -> list[dict[str, object]]:
    """Parse PROJECT:PHASE[:BODY_TYPE_ID]=BODY_TYPE CLI values."""
    parsed: list[dict[str, object]] = []
    for item in items:
        target, separator, body_type = item.partition("=")
        parts = target.split(":")
        if not separator or len(parts) not in {2, 3}:
            raise ValueError(
                f"Invalid rule {item!r}; expected "
                "PROJECT_ID:PHASE_ID[:BODY_TYPE_ID]=BODY_TYPE."
            )
        rule: dict[str, object] = {
            "project_id": parts[0],
            "phase_id": parts[1],
            "body_type": body_type,
        }
        if len(parts) == 3:
            rule["body_type_id"] = parts[2]
        parsed.append(rule)
    return normalize_rules(parsed)


def find_assignments(
    cursor, rules: Sequence[Mapping[str, object]]
) -> list[dict[str, object]]:
    assignments: dict[tuple[int, int], dict[str, object]] = {}
    missing: list[str] = []
    for rule in rules:
        project_id = int(rule["project_id"])
        phase_id = int(rule["phase_id"])
        body_type_name = str(rule["body_type"])
        requested_id = rule.get("body_type_id")
        body_type_id = core.get_unique_body_type_id(
            cursor,
            body_type_name,
            int(requested_id) if requested_id is not None else None,
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
            """,
            (project_id, phase_id, phase_id),
        )
        rows = core.rows_as_dicts(cursor)
        if not rows:
            missing.append(f"project={project_id}:phase={phase_id}")
            continue
        assignment = rows[0]
        assignment.update(
            {
                "phase_id": phase_id,
                "target_body_type": body_type_name,
                "target_body_type_id": body_type_id,
            }
        )
        assignments[(project_id, phase_id)] = assignment
    if missing:
        raise ValueError(f"Project/phase data not found for rules: {missing}")
    return [assignments[key] for key in sorted(assignments)]


def build_plan(
    cursor, rules: Sequence[Mapping[str, object]]
) -> list[dict[str, object]]:
    plan: list[dict[str, object]] = []
    for assignment in find_assignments(cursor, rules):
        target_setup_id = core.find_compatible_setup(cursor, assignment)
        assignment["target_setup_id"] = target_setup_id
        source_counts = core.count_setup_graph(cursor, int(assignment["old_setup_id"]))
        assignment["estimated_inserts"] = core.estimate_insert_counts(
            cursor,
            int(assignment["old_setup_id"]),
            target_setup_id,
            source_counts,
        )
        assignment["estimated_updates"] = core.estimate_update_counts(
            cursor, assignment, target_setup_id
        )
        plan.append(assignment)
    return plan


def print_plan(db_connection, plan: Sequence[Mapping[str, object]]) -> None:
    settings = db_connection.settings_dict
    print(f"Database: {settings.get('NAME')}")
    print(f"Host: {settings.get('HOST')}:{settings.get('PORT') or 3306}")
    print(f"Affected tables: {', '.join(MUTATION_TABLES)}")
    print("Projects selected:")
    for item in plan:
        target_setup = item["target_setup_id"]
        action = f"reuse setup {target_setup}" if target_setup else "clone setup"
        print(
            f"- {item['project_id']} {item['project_code']}, phase={item['phase_id']}: "
            f"{item['old_body_type']} -> {item['target_body_type']} "
            f"(body_type_id={item['target_body_type_id']}); "
            f"old_setup={item['old_setup_id']}; {action}"
        )
        print("  Estimated INSERT rows:")
        for table in INSERT_TABLES:
            print(f"    {table}: {item['estimated_inserts'][table]}")
        print("  Estimated UPDATE rows:")
        for table in UPDATE_TABLES:
            print(f"    {table}: {item['estimated_updates'][table]}")


def plan_fingerprint(plan: Sequence[Mapping[str, object]]) -> list[tuple[object, ...]]:
    return [
        (
            item["project_id"],
            item["project_code"],
            item["phase_id"],
            item["old_setup_id"],
            item["target_setup_id"],
            item["old_body_type_id"],
            item["target_body_type_id"],
            item["destination_id"],
            item["architecture_id"],
            item["drivetrain_id"],
        )
        for item in plan
    ]


def apply_assignment(cursor, assignment: Mapping[str, object]) -> dict[str, int]:
    old_setup_id = int(assignment["old_setup_id"])
    new_setup_id = core.find_compatible_setup(cursor, assignment)
    setup_created = 0
    if new_setup_id is None:
        new_setup_id = core.create_setup(cursor, assignment)
        setup_created = 1
    r_map, r_inserted = core.clone_setup_relations(
        cursor, "setup_r_relation", "r_parameter_id", old_setup_id, new_setup_id
    )
    l_map, l_inserted = core.clone_setup_relations(
        cursor, "setup_l_relation", "l_parameter_id", old_setup_id, new_setup_id
    )
    lr_inserted = core.clone_l_r_relations(cursor, r_map, l_map)
    rfl_map, rfl_inserted = core.clone_rfl_rows(cursor, r_map, l_map)
    rfl_updates = core.update_rfl_references(
        cursor, int(assignment["project_id"]), int(assignment["phase_id"]), rfl_map
    )
    state_updates = core.update_phase_operational_state(cursor, assignment)
    return {
        "setup_inserted": setup_created,
        "setup_r_relation_inserted": r_inserted,
        "setup_l_relation_inserted": l_inserted,
        "l_r_relation_inserted": lr_inserted,
        "rfl_inserted": rfl_inserted,
        **rfl_updates,
        **state_updates,
    }


def print_actual_counts(
    assignment: Mapping[str, object], counts: Mapping[str, int]
) -> None:
    """Print results for the original in-place workflow."""
    print(
        f"\nApplied {assignment['project_id']} {assignment['project_code']}, "
        f"phase={assignment['phase_id']}:"
    )
    print("  INSERT rows:")
    for table in INSERT_TABLES:
        print(f"    {table}: {counts[f'{table}_inserted']}")
    print("  UPDATE rows:")
    for table in UPDATE_TABLES:
        print(f"    {table}: {counts[table]}")


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
    targets = ",".join(f"{x['project_id']}@{x['phase_id']}" for x in plan)
    confirmation = f"UPDATE_BODY_TYPE {database_name}:{targets}"
    if not assume_yes:
        answer = input(f"\nType exactly '{confirmation}' to continue: ").strip()
        if answer != confirmation:
            print("Cancelled. No data was changed.")
            return plan

    lock_name: str | None = None
    try:
        with db_connection.cursor() as cursor:
            lock_name = core.acquire_advisory_lock(cursor, str(database_name))
        with transaction.atomic(using=db_connection.alias):
            with db_connection.cursor() as cursor:
                current_plan = build_plan(cursor, normalized_rules)
                if plan_fingerprint(current_plan) != plan_fingerprint(plan):
                    raise RuntimeError(
                        "Selected project data changed after preview; transaction "
                        "will roll back."
                    )
                for assignment in current_plan:
                    counts = apply_assignment(cursor, assignment)
                    print_actual_counts(assignment, counts)
                for assignment in current_plan:
                    cursor.execute(
                        """
                        SELECT COUNT(*) FROM operational_state
                        WHERE project_id = %s AND phase_id = %s
                          AND NOT (destination_id <=> %s AND body_type_id <=> %s
                                   AND architecture_id <=> %s
                                   AND drivetrain_id <=> %s)
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
                    if int(cursor.fetchone()[0]) != 0:
                        raise RuntimeError(
                            "Operational-state verification failed; transaction "
                            "will roll back."
                        )
        print("\nBody-type update completed successfully.")
        return plan
    finally:
        if lock_name is not None:
            with db_connection.cursor() as cursor:
                core.release_advisory_lock(cursor, lock_name)


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Preview or safely change body types for project/phase IDs."
    )
    parser.add_argument(
        "--rule",
        action="append",
        metavar="PROJECT_ID:PHASE_ID[:BODY_TYPE_ID]=BODY_TYPE",
        help="Repeat for each project/phase pair; overrides RUN rules.",
    )
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--yes", action="store_true")
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    rules = parse_rule_items(args.rule) if args.rule else RUN_PROJECT_BODY_TYPE_RULES
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES
    if assume_yes and not execute:
        raise ValueError("RUN_ASSUME_YES/--yes requires execute mode.")
    edit_project_body_types(rules, execute=execute, assume_yes=assume_yes)


if __name__ == "__main__":
    main()
