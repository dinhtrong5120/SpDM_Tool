"""Preview or remap old R parameter IDs to one existing R parameter ID.

The tool updates references; it deliberately does not delete rows from
``r_parameter``. For PyCharm's Run button, edit the RUN_* constants below.
The default is read-only preview mode.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402


# PyCharm Run-button configuration. Preview first, then explicitly enable run.
RUN_OLD_R_PARAMETER_IDS = [290938]
RUN_NEW_R_PARAMETER_ID = 70056
RUN_EXECUTE = False
RUN_ASSUME_YES = False

IDENTIFIER_PATTERN = re.compile(r"^[A-Za-z0-9_$]+$")
LEGACY_R_PARAMETER_COLUMNS = frozenset(
    {
        ("detail_data_rlist", "r_parameter_id"),
    }
)
REFERENCE_EXCLUSIONS = {
    "r_parameter": frozenset({("setup_r_relation", "r_parameter_id")}),
    "rfl": frozenset({("prj_rfl", "rfl_id")}),
    "setup_r_relation": frozenset(
        {
            ("rfl", "requirement_s_id"),
            ("l_r_relation", "r_s_id"),
        }
    ),
}


@dataclass(frozen=True, order=True)
class Reference:
    table: str
    column: str


@dataclass(frozen=True)
class ReferencePlan:
    reference: Reference
    mappings: tuple[tuple[int, int], ...]
    row_counts: tuple[int, ...]
    unique_conflicts: tuple[str, ...] = ()

    @property
    def row_count(self) -> int:
        return sum(self.row_counts)


@dataclass(frozen=True)
class MasterListChange:
    row_id: int
    old_json: str
    new_json: str


@dataclass(frozen=True)
class MigrationPlan:
    old_ids: tuple[int, ...]
    new_id: int
    parameter_rows: tuple[tuple[Any, ...], ...]
    r_parameter_references: tuple[Reference, ...]
    direct_reference_plans: tuple[ReferencePlan, ...]
    setup_relation_updates: tuple[int, ...]
    setup_relation_merges: tuple[tuple[int, int], ...]
    rfl_updates: tuple[tuple[int, int, int], ...]
    rfl_merges: tuple[tuple[int, int], ...]
    rfl_reference_plans: tuple[ReferencePlan, ...]
    prj_rfl_deletes: tuple[tuple[int, int, int], ...]
    prj_rfl_updates: tuple[tuple[int, int, int, int], ...]
    setup_reference_plans: tuple[ReferencePlan, ...]
    l_r_relation_deletes: tuple[int, ...]
    l_r_relation_updates: tuple[tuple[int, int, int], ...]
    master_list_changes: tuple[MasterListChange, ...]
    conflicts: tuple[str, ...]

    @property
    def mutation_count(self) -> int:
        return (
            sum(plan.row_count for plan in self.direct_reference_plans)
            + len(self.setup_relation_updates)
            + len(self.setup_relation_merges)
            + len(self.rfl_updates)
            + len(self.rfl_merges)
            + sum(plan.row_count for plan in self.rfl_reference_plans)
            + len(self.prj_rfl_deletes)
            + len(self.prj_rfl_updates)
            + sum(plan.row_count for plan in self.setup_reference_plans)
            + len(self.l_r_relation_deletes)
            + len(self.l_r_relation_updates)
            + len(self.master_list_changes)
        )


def quote_identifier(identifier: str) -> str:
    if not IDENTIFIER_PATTERN.fullmatch(identifier):
        raise ValueError(f"Invalid SQL identifier: {identifier!r}")
    return f"`{identifier}`"


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create placeholders for an empty sequence.")
    return ", ".join(["%s"] * len(values))


def normalize_positive_id(raw_id: object, *, label: str) -> int:
    if isinstance(raw_id, bool):
        raise ValueError(f"{label} must be a positive integer, not boolean.")
    if isinstance(raw_id, int):
        normalized = raw_id
    elif isinstance(raw_id, str) and re.fullmatch(r"\+?\d+", raw_id.strip()):
        normalized = int(raw_id.strip())
    else:
        raise ValueError(f"Invalid {label}: {raw_id!r}")
    if normalized <= 0:
        raise ValueError(f"Invalid {label}: {raw_id!r}")
    return normalized


def normalize_id_mapping(
    old_ids: Iterable[object], new_id: object
) -> tuple[tuple[int, ...], int]:
    normalized_new = normalize_positive_id(new_id, label="new R parameter ID")
    normalized_old: list[int] = []
    seen: set[int] = set()
    for index, raw_id in enumerate(old_ids, start=1):
        old_id = normalize_positive_id(raw_id, label=f"old R parameter ID #{index}")
        if old_id == normalized_new:
            raise ValueError("The new R parameter ID cannot also be an old ID.")
        if old_id not in seen:
            seen.add(old_id)
            normalized_old.append(old_id)
    if not normalized_old:
        raise ValueError("At least one old R parameter ID is required.")
    return tuple(normalized_old), normalized_new


def replace_integer_ids(value: Any, id_map: Mapping[int, int]) -> Any:
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return id_map.get(value, value)
    if isinstance(value, list):
        return [replace_integer_ids(item, id_map) for item in value]
    if isinstance(value, dict):
        return {key: replace_integer_ids(item, id_map) for key, item in value.items()}
    return value


def database_identity() -> tuple[str, str]:
    settings_dict = connection.settings_dict
    return (
        str(settings_dict.get("NAME") or "<unknown>"),
        str(settings_dict.get("HOST") or "localhost"),
    )


def discover_references(cursor: Any, referenced_table: str) -> tuple[Reference, ...]:
    cursor.execute(
        """
        SELECT DISTINCT kcu.TABLE_NAME, kcu.COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE AS kcu
        JOIN information_schema.TABLES AS tables_meta
          ON tables_meta.TABLE_SCHEMA = kcu.TABLE_SCHEMA
         AND tables_meta.TABLE_NAME = kcu.TABLE_NAME
         AND tables_meta.TABLE_TYPE = 'BASE TABLE'
        WHERE kcu.TABLE_SCHEMA = DATABASE()
          AND kcu.REFERENCED_TABLE_NAME = %s
          AND kcu.REFERENCED_COLUMN_NAME = 'id'
        ORDER BY kcu.TABLE_NAME, kcu.COLUMN_NAME
        """,
        [referenced_table],
    )
    return tuple(Reference(str(row[0]), str(row[1])) for row in cursor.fetchall())


def column_exists(cursor: Any, reference: Reference) -> bool:
    cursor.execute(
        """
        SELECT COUNT(*)
        FROM information_schema.COLUMNS AS columns_meta
        JOIN information_schema.TABLES AS tables_meta
          ON tables_meta.TABLE_SCHEMA = columns_meta.TABLE_SCHEMA
         AND tables_meta.TABLE_NAME = columns_meta.TABLE_NAME
         AND tables_meta.TABLE_TYPE = 'BASE TABLE'
        WHERE columns_meta.TABLE_SCHEMA = DATABASE()
          AND columns_meta.TABLE_NAME = %s
          AND columns_meta.COLUMN_NAME = %s
        """,
        [reference.table, reference.column],
    )
    return int(cursor.fetchone()[0]) == 1


def discover_unique_indexes(
    cursor: Any, reference: Reference
) -> tuple[tuple[str, tuple[str, ...]], ...]:
    cursor.execute(
        """
        SELECT INDEX_NAME, COLUMN_NAME
        FROM information_schema.STATISTICS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = %s
          AND NON_UNIQUE = 0
        ORDER BY INDEX_NAME, SEQ_IN_INDEX
        """,
        [reference.table],
    )
    indexes: dict[str, list[str]] = {}
    for index_name, column_name in cursor.fetchall():
        indexes.setdefault(str(index_name), []).append(str(column_name))
    return tuple(
        (index_name, tuple(columns))
        for index_name, columns in sorted(indexes.items())
        if reference.column in columns
    )


def select_reference_values(
    cursor: Any,
    reference: Reference,
    values: Sequence[int],
    *,
    extra_columns: Sequence[str] = (),
    lock_rows: bool,
) -> list[tuple[Any, ...]]:
    placeholders = make_placeholders(values)
    table = quote_identifier(reference.table)
    column = quote_identifier(reference.column)
    selected_columns = [column, *map(quote_identifier, extra_columns)]
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"SELECT {', '.join(selected_columns)} FROM {table} "
        f"WHERE {column} IN ({placeholders}){lock_clause}",
        list(values),
    )
    return list(cursor.fetchall())


def find_unique_conflicts(
    cursor: Any,
    reference: Reference,
    mappings: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
) -> tuple[str, ...]:
    mapping = dict(mappings)
    relevant_values = tuple(dict.fromkeys([*mapping, *mapping.values()]))
    conflicts: list[str] = []
    for index_name, index_columns in discover_unique_indexes(cursor, reference):
        other_columns = tuple(
            column for column in index_columns if column != reference.column
        )
        rows = select_reference_values(
            cursor,
            reference,
            relevant_values,
            extra_columns=other_columns,
            lock_rows=lock_rows,
        )
        seen: set[tuple[Any, ...]] = set()
        conflict_count = 0
        for row in rows:
            remapped_key = (mapping.get(int(row[0]), int(row[0])), *row[1:])
            # MySQL unique indexes permit multiple rows when any component is NULL.
            if any(value is None for value in remapped_key):
                continue
            if remapped_key in seen:
                conflict_count += 1
            else:
                seen.add(remapped_key)
        if conflict_count:
            conflicts.append(
                f"{reference.table}.{index_name}: {conflict_count} collision(s)"
            )
    return tuple(conflicts)


def build_reference_plan(
    cursor: Any,
    reference: Reference,
    mappings: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
    check_unique: bool = True,
) -> ReferencePlan:
    table = quote_identifier(reference.table)
    column = quote_identifier(reference.column)
    row_counts: list[int] = []
    lock_clause = " FOR UPDATE" if lock_rows else ""
    for old_id, _new_id in mappings:
        if lock_rows:
            cursor.execute(
                f"SELECT {column} FROM {table} WHERE {column} = %s{lock_clause}",
                [old_id],
            )
            row_counts.append(len(cursor.fetchall()))
        else:
            cursor.execute(
                f"SELECT COUNT(*) FROM {table} WHERE {column} = %s",
                [old_id],
            )
            row_counts.append(int(cursor.fetchone()[0]))
    unique_conflicts = (
        find_unique_conflicts(cursor, reference, mappings, lock_rows=lock_rows)
        if check_unique and any(row_counts)
        else ()
    )
    return ReferencePlan(
        reference=reference,
        mappings=tuple(mappings),
        row_counts=tuple(row_counts),
        unique_conflicts=unique_conflicts,
    )


def build_setup_relation_plan(
    cursor: Any,
    old_ids: Sequence[int],
    new_id: int,
    *,
    lock_rows: bool,
) -> tuple[tuple[int, ...], tuple[tuple[int, int], ...]]:
    values = [*old_ids, new_id]
    placeholders = make_placeholders(values)
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"""
        SELECT id, setup_id, r_parameter_id, usecase_id
        FROM setup_r_relation
        WHERE r_parameter_id IN ({placeholders})
        ORDER BY id{lock_clause}
        """,
        values,
    )
    rows = [tuple(map(int, row)) for row in cursor.fetchall()]
    targets = {
        (setup_id, usecase_id): relation_id
        for relation_id, setup_id, parameter_id, usecase_id in rows
        if parameter_id == new_id
    }
    updates: list[int] = []
    merges: list[tuple[int, int]] = []
    for relation_id, setup_id, parameter_id, usecase_id in rows:
        if parameter_id not in old_ids:
            continue
        key = (setup_id, usecase_id)
        target_id = targets.get(key)
        if target_id is None:
            updates.append(relation_id)
            targets[key] = relation_id
        else:
            merges.append((relation_id, target_id))
    return tuple(updates), tuple(merges)


def build_rfl_plan(
    cursor: Any,
    setup_merges: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
) -> tuple[tuple[tuple[int, int, int], ...], tuple[tuple[int, int], ...]]:
    if not setup_merges:
        return (), ()
    relation_ids = tuple(
        dict.fromkeys(value for pair in setup_merges for value in pair)
    )
    placeholders = make_placeholders(relation_ids)
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"""
        SELECT id, requirement_s_id, function_id, logic_s_id
        FROM rfl
        WHERE requirement_s_id IN ({placeholders})
        ORDER BY id{lock_clause}
        """,
        list(relation_ids),
    )
    rows = []
    for raw_id, raw_relation_id, raw_function_id, raw_logic_id in cursor.fetchall():
        if raw_id is None:
            raise ValueError(
                "A selected rfl row has NULL id; it cannot be merged safely."
            )
        rows.append(
            (
                int(raw_id),
                int(raw_relation_id),
                int(raw_function_id),
                int(raw_logic_id),
            )
        )
    rows_by_relation: dict[int, list[tuple[int, int, int]]] = {}
    for rfl_id, relation_id, function_id, logic_id in rows:
        rows_by_relation.setdefault(relation_id, []).append(
            (rfl_id, function_id, logic_id)
        )

    target_rows: dict[tuple[int, int, int], int] = {}
    for _source_id, target_id in setup_merges:
        for rfl_id, function_id, logic_id in rows_by_relation.get(target_id, []):
            target_rows[(target_id, function_id, logic_id)] = rfl_id

    updates: list[tuple[int, int, int]] = []
    merges: list[tuple[int, int]] = []
    for source_id, target_id in setup_merges:
        for rfl_id, function_id, logic_id in rows_by_relation.get(source_id, []):
            key = (target_id, function_id, logic_id)
            target_rfl_id = target_rows.get(key)
            if target_rfl_id is None:
                updates.append((rfl_id, source_id, target_id))
                target_rows[key] = rfl_id
            else:
                merges.append((rfl_id, target_rfl_id))
    return tuple(updates), tuple(merges)


def build_prj_rfl_plan(
    cursor: Any,
    rfl_merges: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
) -> tuple[tuple[tuple[int, int, int], ...], tuple[tuple[int, int, int, int], ...]]:
    if not rfl_merges:
        return (), ()
    rfl_ids = tuple(dict.fromkeys(value for pair in rfl_merges for value in pair))
    placeholders = make_placeholders(rfl_ids)
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"""
        SELECT project_info_id, rfl_id, phase_id
        FROM prj_rfl
        WHERE rfl_id IN ({placeholders})
        ORDER BY project_info_id, phase_id, rfl_id{lock_clause}
        """,
        list(rfl_ids),
    )
    active_rows = {tuple(map(int, row)) for row in cursor.fetchall()}
    deletes: list[tuple[int, int, int]] = []
    updates: list[tuple[int, int, int, int]] = []
    for old_rfl_id, new_rfl_id in rfl_merges:
        old_rows = sorted(row for row in active_rows if row[1] == old_rfl_id)
        for project_id, _old_id, phase_id in old_rows:
            old_key = (project_id, old_rfl_id, phase_id)
            new_key = (project_id, new_rfl_id, phase_id)
            if new_key in active_rows:
                deletes.append(old_key)
                active_rows.remove(old_key)
            else:
                updates.append((project_id, phase_id, old_rfl_id, new_rfl_id))
                active_rows.remove(old_key)
                active_rows.add(new_key)
    return tuple(deletes), tuple(updates)


def build_l_r_relation_plan(
    cursor: Any,
    setup_merges: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
) -> tuple[tuple[int, ...], tuple[tuple[int, int, int], ...]]:
    if not setup_merges:
        return (), ()
    relation_ids = tuple(
        dict.fromkeys(value for pair in setup_merges for value in pair)
    )
    placeholders = make_placeholders(relation_ids)
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"""
        SELECT id, r_s_id, l_s_id
        FROM l_r_relation
        WHERE r_s_id IN ({placeholders})
        ORDER BY id{lock_clause}
        """,
        list(relation_ids),
    )
    rows = [
        (int(row_id), int(r_s_id), int(l_s_id) if l_s_id is not None else None)
        for row_id, r_s_id, l_s_id in cursor.fetchall()
    ]
    active_keys = {(r_s_id, l_s_id): row_id for row_id, r_s_id, l_s_id in rows}
    rows_by_relation: dict[int, list[tuple[int, int]]] = {}
    for row_id, r_s_id, l_s_id in rows:
        rows_by_relation.setdefault(r_s_id, []).append((row_id, l_s_id))

    deletes: list[int] = []
    updates: list[tuple[int, int, int]] = []
    for source_id, target_id in setup_merges:
        for row_id, l_s_id in rows_by_relation.get(source_id, []):
            if l_s_id is not None and (target_id, l_s_id) in active_keys:
                deletes.append(row_id)
            else:
                updates.append((row_id, source_id, target_id))
                if l_s_id is not None:
                    active_keys[(target_id, l_s_id)] = row_id
    return tuple(deletes), tuple(updates)


def build_master_list_changes(
    cursor: Any,
    rfl_merges: Sequence[tuple[int, int]],
    *,
    lock_rows: bool,
) -> tuple[MasterListChange, ...]:
    if not rfl_merges:
        return ()
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"SELECT id, list_order FROM master_list_order ORDER BY id{lock_clause}"
    )
    id_map = dict(rfl_merges)
    changes: list[MasterListChange] = []
    for row_id, raw_value in cursor.fetchall():
        parsed = json.loads(raw_value) if isinstance(raw_value, str) else raw_value
        replaced = replace_integer_ids(parsed, id_map)
        if replaced == parsed:
            continue
        old_json = json.dumps(parsed, ensure_ascii=False, separators=(",", ":"))
        new_json = json.dumps(replaced, ensure_ascii=False, separators=(",", ":"))
        changes.append(MasterListChange(int(row_id), old_json, new_json))
    return tuple(changes)


def build_plan(
    cursor: Any,
    old_ids: Sequence[int],
    new_id: int,
    *,
    lock_rows: bool,
) -> MigrationPlan:
    all_parameter_ids = [*old_ids, new_id]
    placeholders = make_placeholders(all_parameter_ids)
    lock_clause = " FOR UPDATE" if lock_rows else ""
    cursor.execute(
        f"""
        SELECT id, performance, design_item_1, design_item_2, design_item_3,
               unit, hierarchy_id, wp_id, unit_id
        FROM r_parameter
        WHERE id IN ({placeholders})
        ORDER BY id{lock_clause}
        """,
        all_parameter_ids,
    )
    parameter_rows = tuple(cursor.fetchall())
    found_ids = {int(row[0]) for row in parameter_rows}
    missing_ids = [value for value in all_parameter_ids if value not in found_ids]
    if missing_ids:
        raise ValueError(f"R parameter IDs do not exist: {missing_ids}")

    r_parameter_references = list(discover_references(cursor, "r_parameter"))
    for table_name, column_name in sorted(LEGACY_R_PARAMETER_COLUMNS):
        legacy_reference = Reference(table_name, column_name)
        if (
            column_exists(cursor, legacy_reference)
            and legacy_reference not in r_parameter_references
        ):
            r_parameter_references.append(legacy_reference)
    r_parameter_references.sort()

    setup_updates, setup_merges = build_setup_relation_plan(
        cursor, old_ids, new_id, lock_rows=lock_rows
    )
    rfl_updates, rfl_merges = build_rfl_plan(cursor, setup_merges, lock_rows=lock_rows)
    prj_deletes, prj_updates = build_prj_rfl_plan(
        cursor, rfl_merges, lock_rows=lock_rows
    )
    lrr_deletes, lrr_updates = build_l_r_relation_plan(
        cursor, setup_merges, lock_rows=lock_rows
    )
    master_changes = build_master_list_changes(cursor, rfl_merges, lock_rows=lock_rows)

    direct_mapping = tuple((old_id, new_id) for old_id in old_ids)
    direct_plans: list[ReferencePlan] = []
    for reference in r_parameter_references:
        if (reference.table, reference.column) in REFERENCE_EXCLUSIONS["r_parameter"]:
            continue
        direct_plans.append(
            build_reference_plan(
                cursor,
                reference,
                direct_mapping,
                lock_rows=lock_rows,
            )
        )

    rfl_refs = discover_references(cursor, "rfl")
    rfl_ref_plans: list[ReferencePlan] = []
    for reference in rfl_refs:
        if (reference.table, reference.column) in REFERENCE_EXCLUSIONS["rfl"]:
            continue
        rfl_ref_plans.append(
            build_reference_plan(
                cursor,
                reference,
                rfl_merges,
                lock_rows=lock_rows,
            )
        )

    setup_refs = discover_references(cursor, "setup_r_relation")
    setup_ref_plans: list[ReferencePlan] = []
    for reference in setup_refs:
        if (reference.table, reference.column) in REFERENCE_EXCLUSIONS[
            "setup_r_relation"
        ]:
            continue
        setup_ref_plans.append(
            build_reference_plan(
                cursor,
                reference,
                setup_merges,
                lock_rows=lock_rows,
            )
        )

    conflicts = tuple(
        conflict
        for plan in (*direct_plans, *rfl_ref_plans, *setup_ref_plans)
        for conflict in plan.unique_conflicts
    )
    return MigrationPlan(
        old_ids=tuple(old_ids),
        new_id=new_id,
        parameter_rows=parameter_rows,
        r_parameter_references=tuple(r_parameter_references),
        direct_reference_plans=tuple(direct_plans),
        setup_relation_updates=setup_updates,
        setup_relation_merges=setup_merges,
        rfl_updates=rfl_updates,
        rfl_merges=rfl_merges,
        rfl_reference_plans=tuple(rfl_ref_plans),
        prj_rfl_deletes=prj_deletes,
        prj_rfl_updates=prj_updates,
        setup_reference_plans=tuple(setup_ref_plans),
        l_r_relation_deletes=lrr_deletes,
        l_r_relation_updates=lrr_updates,
        master_list_changes=master_changes,
        conflicts=conflicts,
    )


def print_plan(plan: MigrationPlan, *, database: str, host: str) -> None:
    print("R parameter ID remap preview")
    print(f"Database: {database}")
    print(f"Host: {host}")
    print(f"Old IDs: {', '.join(map(str, plan.old_ids))}")
    print(f"New ID: {plan.new_id}")
    print("Selected r_parameter records:")
    for row in plan.parameter_rows:
        marker = "TARGET" if int(row[0]) == plan.new_id else "SOURCE"
        print(f"  [{marker}] {row}")
    print("Affected tables:")
    for reference_plan in (
        *plan.direct_reference_plans,
        *plan.rfl_reference_plans,
        *plan.setup_reference_plans,
    ):
        if reference_plan.row_count:
            reference = reference_plan.reference
            print(
                f"  {reference.table}.{reference.column}: "
                f"{reference_plan.row_count} row(s)"
            )
    print(f"  setup_r_relation rows updated: {len(plan.setup_relation_updates)}")
    print(f"  setup_r_relation rows merged/deleted: {len(plan.setup_relation_merges)}")
    print(f"  rfl rows rewired: {len(plan.rfl_updates)}")
    print(f"  rfl rows merged/deleted: {len(plan.rfl_merges)}")
    print(f"  prj_rfl duplicate rows deleted: {len(plan.prj_rfl_deletes)}")
    print(f"  prj_rfl rows updated: {len(plan.prj_rfl_updates)}")
    print(f"  l_r_relation duplicate rows deleted: {len(plan.l_r_relation_deletes)}")
    print(f"  l_r_relation rows updated: {len(plan.l_r_relation_updates)}")
    print(f"  master_list_order rows updated: {len(plan.master_list_changes)}")
    print(f"Estimated mutation statements/rows: {plan.mutation_count}")
    if plan.setup_relation_updates:
        print(f"setup_r_relation IDs to update: {plan.setup_relation_updates}")
    if plan.setup_relation_merges:
        print(f"setup_r_relation merge map: {plan.setup_relation_merges}")
    if plan.rfl_updates:
        print(
            f"rfl rows to rewire (id, old relation, new relation): {plan.rfl_updates}"
        )
    if plan.rfl_merges:
        print(f"rfl merge map: {plan.rfl_merges}")
    if plan.prj_rfl_deletes:
        print(f"prj_rfl keys to delete: {plan.prj_rfl_deletes}")
    if plan.l_r_relation_deletes:
        print(f"l_r_relation IDs to delete: {plan.l_r_relation_deletes}")
    if plan.conflicts:
        print("BLOCKING unique-key conflicts:")
        for conflict in plan.conflicts:
            print(f"  - {conflict}")
    elif plan.mutation_count == 0:
        print("No references require changes; this is a successful no-op.")


def require_confirmation(database: str, old_ids: Sequence[int], new_id: int) -> None:
    expected = f"REMAP {database} {','.join(map(str, old_ids))} TO {new_id}"
    print("\nMutation requested. Type the following text exactly to continue:")
    print(expected)
    if input("> ").strip() != expected:
        raise RuntimeError("Confirmation did not match; nothing was changed.")


def execute_reference_plan(cursor: Any, plan: ReferencePlan) -> None:
    table = quote_identifier(plan.reference.table)
    column = quote_identifier(plan.reference.column)
    for (old_id, new_id), expected_count in zip(
        plan.mappings, plan.row_counts, strict=True
    ):
        if expected_count == 0:
            continue
        cursor.execute(
            f"UPDATE {table} SET {column} = %s WHERE {column} = %s",
            [new_id, old_id],
        )
        if cursor.rowcount != expected_count:
            raise RuntimeError(
                f"Unexpected update count for {plan.reference.table}."
                f"{plan.reference.column}: got {cursor.rowcount}, "
                f"expected {expected_count}."
            )


def expect_row_count(cursor: Any, expected: int, action: str) -> None:
    if cursor.rowcount != expected:
        raise RuntimeError(
            f"Unexpected row count for {action}: got {cursor.rowcount}, "
            f"expected {expected}."
        )


def execute_plan(cursor: Any, plan: MigrationPlan) -> None:
    if plan.conflicts:
        raise RuntimeError("Cannot execute a plan with unique-key conflicts.")

    for reference_plan in plan.rfl_reference_plans:
        execute_reference_plan(cursor, reference_plan)

    for project_id, old_rfl_id, phase_id in plan.prj_rfl_deletes:
        cursor.execute(
            """
            DELETE FROM prj_rfl
            WHERE project_info_id = %s AND rfl_id = %s AND phase_id = %s
            """,
            [project_id, old_rfl_id, phase_id],
        )
        expect_row_count(cursor, 1, "delete duplicate prj_rfl")
    for project_id, phase_id, old_rfl_id, new_rfl_id in plan.prj_rfl_updates:
        cursor.execute(
            """
            UPDATE prj_rfl SET rfl_id = %s
            WHERE project_info_id = %s AND phase_id = %s AND rfl_id = %s
            """,
            [new_rfl_id, project_id, phase_id, old_rfl_id],
        )
        expect_row_count(cursor, 1, "update prj_rfl.rfl_id")

    for change in plan.master_list_changes:
        cursor.execute(
            "UPDATE master_list_order SET list_order = %s WHERE id = %s",
            [change.new_json, change.row_id],
        )
        expect_row_count(cursor, 1, "update master_list_order.list_order")

    for old_rfl_id, _new_rfl_id in plan.rfl_merges:
        cursor.execute("DELETE FROM rfl WHERE id = %s", [old_rfl_id])
        expect_row_count(cursor, 1, "delete duplicate rfl")
    for rfl_id, old_relation_id, new_relation_id in plan.rfl_updates:
        cursor.execute(
            """
            UPDATE rfl SET requirement_s_id = %s
            WHERE id = %s AND requirement_s_id = %s
            """,
            [new_relation_id, rfl_id, old_relation_id],
        )
        expect_row_count(cursor, 1, "update rfl.requirement_s_id")

    for row_id in plan.l_r_relation_deletes:
        cursor.execute("DELETE FROM l_r_relation WHERE id = %s", [row_id])
        expect_row_count(cursor, 1, "delete duplicate l_r_relation")
    for row_id, old_relation_id, new_relation_id in plan.l_r_relation_updates:
        cursor.execute(
            """
            UPDATE l_r_relation SET r_s_id = %s
            WHERE id = %s AND r_s_id = %s
            """,
            [new_relation_id, row_id, old_relation_id],
        )
        expect_row_count(cursor, 1, "update l_r_relation.r_s_id")
    for reference_plan in plan.setup_reference_plans:
        execute_reference_plan(cursor, reference_plan)

    for old_relation_id, _new_relation_id in plan.setup_relation_merges:
        cursor.execute("DELETE FROM setup_r_relation WHERE id = %s", [old_relation_id])
        expect_row_count(cursor, 1, "delete duplicate setup_r_relation")
    if plan.setup_relation_updates:
        placeholders = make_placeholders(plan.setup_relation_updates)
        cursor.execute(
            f"""
            UPDATE setup_r_relation SET r_parameter_id = %s
            WHERE id IN ({placeholders})
            """,
            [plan.new_id, *plan.setup_relation_updates],
        )
        expect_row_count(
            cursor, len(plan.setup_relation_updates), "update setup_r_relation"
        )

    for reference_plan in plan.direct_reference_plans:
        execute_reference_plan(cursor, reference_plan)


def count_remaining_values(
    cursor: Any, reference: Reference, values: Sequence[int]
) -> int:
    placeholders = make_placeholders(values)
    table = quote_identifier(reference.table)
    column = quote_identifier(reference.column)
    cursor.execute(
        f"SELECT COUNT(*) FROM {table} WHERE {column} IN ({placeholders})",
        list(values),
    )
    return int(cursor.fetchone()[0])


def verify_result(cursor: Any, plan: MigrationPlan) -> None:
    for reference in plan.r_parameter_references:
        remaining = count_remaining_values(cursor, reference, plan.old_ids)
        if remaining:
            raise RuntimeError(
                f"Post-update verification failed: {reference.table}."
                f"{reference.column} still has {remaining} old reference(s)."
            )

    old_rfl_ids = tuple(old_id for old_id, _new_id in plan.rfl_merges)
    if old_rfl_ids:
        cursor.execute(
            f"SELECT COUNT(*) FROM rfl WHERE id IN ({make_placeholders(old_rfl_ids)})",
            list(old_rfl_ids),
        )
        if int(cursor.fetchone()[0]):
            raise RuntimeError("Post-update verification found duplicate old rfl rows.")

    old_relation_ids = tuple(old_id for old_id, _new_id in plan.setup_relation_merges)
    if old_relation_ids:
        cursor.execute(
            "SELECT COUNT(*) FROM setup_r_relation "
            f"WHERE id IN ({make_placeholders(old_relation_ids)})",
            list(old_relation_ids),
        )
        if int(cursor.fetchone()[0]):
            raise RuntimeError(
                "Post-update verification found duplicate old setup_r_relation rows."
            )


def remap_r_parameter_ids(
    old_ids: Iterable[object],
    new_id: object,
    *,
    execute: bool = False,
    assume_yes: bool = False,
) -> MigrationPlan:
    normalized_old_ids, normalized_new_id = normalize_id_mapping(old_ids, new_id)
    database, host = database_identity()
    with connection.cursor() as cursor:
        preview_plan = build_plan(
            cursor,
            normalized_old_ids,
            normalized_new_id,
            lock_rows=False,
        )
    print_plan(preview_plan, database=database, host=host)

    if preview_plan.conflicts:
        raise RuntimeError(
            "Preview found unique-key conflicts. Resolve them explicitly before remapping."
        )
    if not execute:
        print("\nPREVIEW ONLY: no database rows were changed.")
        return preview_plan
    if preview_plan.mutation_count == 0:
        print("\nNothing to update: all references already use the target ID.")
        return preview_plan
    if not assume_yes:
        require_confirmation(database, normalized_old_ids, normalized_new_id)

    with transaction.atomic():
        with connection.cursor() as cursor:
            locked_plan = build_plan(
                cursor,
                normalized_old_ids,
                normalized_new_id,
                lock_rows=True,
            )
            if locked_plan != preview_plan:
                raise RuntimeError(
                    "Database state changed after preview; transaction rolled back. "
                    "Run preview again."
                )
            execute_plan(cursor, locked_plan)
            verify_result(cursor, locked_plan)

    print(f"\nCommitted R parameter ID remap in database {database}.")
    return preview_plan


def parse_id_list(raw_value: str) -> tuple[int, ...]:
    parts = [part.strip() for part in raw_value.split(",") if part.strip()]
    normalized: list[int] = []
    seen: set[int] = set()
    for index, part in enumerate(parts, start=1):
        value = normalize_positive_id(part, label=f"old R parameter ID #{index}")
        if value not in seen:
            seen.add(value)
            normalized.append(value)
    if not normalized:
        raise ValueError("At least one old R parameter ID is required.")
    return tuple(normalized)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preview or remap old r_parameter IDs to one existing ID."
    )
    parser.add_argument(
        "--old-ids",
        default=",".join(map(str, RUN_OLD_R_PARAMETER_IDS)),
        help="Comma-separated old r_parameter IDs.",
    )
    parser.add_argument(
        "--new-id",
        type=int,
        default=RUN_NEW_R_PARAMETER_ID,
        help="Existing target r_parameter ID.",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Apply the remap; without this flag the tool only previews.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip confirmation (only for audited non-interactive runs).",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    remap_r_parameter_ids(
        parse_id_list(args.old_ids),
        args.new_id,
        execute=args.execute or RUN_EXECUTE,
        assume_yes=args.yes or RUN_ASSUME_YES,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
