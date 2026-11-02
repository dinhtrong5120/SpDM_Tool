"""Copy one project/phase between MySQL schemas. See copy_project_from_old.md.

No database access at import time or with --help. Run while application writers
are stopped: legacy relationships are not protected by database constraints.
"""

import argparse
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path

RUN_PROJECT_ID = 1054
RUN_PHASE_ID = 16
RUN_SOURCE_DATABASE = "dev_spdm"
RUN_TARGET_DATABASE = "spdm"
RUN_EXECUTE = False
RUN_REPORT_DIR = Path(__file__).resolve().parents[2] / "outputs" / "project_copy"

# Reviewed against project/views.py, utils/{project,se,rlist,scenario_config}_module.py,
# migrations and utils/db_scripts/view/*.sql. Unknown project columns block the run.
LEGACY_PROJECT_TABLES = {
    "project_parameter",
    "project_r_parameter",
    "detail_data",
    "detail_data_rlist",
    "detail_data_design",
    "test_project_senario_parameter",
    "project_statement",
    "project_parameter_attachments",
    "parameter_map_relation",
    "prj_rfl",
    "project_scenario_parameter",
    "project_r_parameter_history",
    "project_se_notification",
    "project_se_notification_target",
    "project_se_input_completion",
    "project_se_confirmation_history",
    "study_id_note",
    "study_group",
}
# (child, child columns, parent, parent columns, copy children with selected parent)
LEGACY_RELATIONS = [
    ("comments_table", ("statement_id",), "study_group_statement", ("id",), True),
    ("se_parameter_target_list", ("se_parameter_id",), "se_parameter", ("id",), True),
    ("se_parameter_usage_domain", ("se_parameter_id",), "se_parameter", ("id",), True),
    ("se_parameter_target_list", ("target_list_id",), "target_list", ("id",), False),
    ("se_parameter_usage_domain", ("usage_domain_id",), "usage_domain", ("id",), False),
    ("study_id_links", ("study_links_note",), "study_id_note", ("id",), True),
    ("study_group_statement", ("study_group_id",), "study_group", ("id",), True),
    ("statement_history", ("study_group_id",), "study_group", ("id",), True),
    ("statement_history", ("statement_id",), "study_group_statement", ("id",), False),
    (
        "project_parameter_attachments",
        ("type_id",),
        "project_parameter_attachment_types",
        ("id",),
        False,
    ),
    ("detail_data_sim", ("id",), "test_project_senario_parameter", ("id",), True),
    (
        "project_se_notification_target",
        ("notification_id",),
        "project_se_notification",
        ("id",),
        False,
    ),
    ("prj_rfl", ("rfl_id",), "rfl", ("id",), False),
    ("rfl", ("requirement_s_id",), "setup_r_relation", ("id",), False),
    ("rfl", ("logic_s_id",), "setup_l_relation", ("id",), False),
    ("rfl", ("function_id",), "f_parameter", ("id",), False),
    ("l_r_relation", ("l_s_id",), "setup_l_relation", ("id",), True),
    ("l_r_relation", ("r_s_id",), "setup_r_relation", ("id",), False),
    ("usecase_detail", ("usecase_id",), "usecase", ("id",), True),
]
LEGACY_COLUMNS = {
    "project_id": "project_info",
    "project_info_id": "project_info",
    "phase_id": "phase",
    "variation_id": "variation",
    "setup_id": "setup",
    "lot_id": "lot",
    "family_id": "family",
    "architecture_id": "architecture",
    "destination_id": "destination",
    "drivetrain_id": "drivetrain",
    "body_type_id": "body_type",
    "se_parameter_id": "se_parameter",
    "r_parameter_id": "r_parameter",
    "f_parameter_id": "f_parameter",
    "l_parameter_id": "l_parameter",
    "wp_id": "wp",
    "unit_id": "units",
    "hierarchy_id": "hierarchy",
    "usecase_id": "usecase",
    "usecase_parameter_id": "usecase_parameter",
    "map_variable_id": "map_variable",
    "senario_parameter_id": "test_senario_list_parameter",
    "scenario_parameter_id": "test_senario_list_parameter",
}
# Only apply the column map to these reviewed tables, never arbitrary *_id columns.
LEGACY_COLUMN_TABLES = LEGACY_PROJECT_TABLES | {
    "project_info",
    "setup",
    "se_parameter",
    "r_parameter",
    "f_parameter",
    "l_parameter",
    "setup_r_relation",
    "setup_l_relation",
    "usecase_detail",
    "map_structure",
    "detail_data_map_name",
    "detail_data_map",
    "test_senario_list_parameter",
}


def ident(value):
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", value):
        raise ValueError(f"Invalid SQL identifier: {value!r}")
    return f"`{value}`"


def positive_id(value):
    if (
        isinstance(value, bool)
        or not re.fullmatch(r"[0-9]+", str(value))
        or int(value) < 1
    ):
        raise ValueError(f"ID must be a positive integer: {value!r}")
    return int(value)


def predicate(columns, values, json_columns=frozenset()):
    if not columns or len(columns) != len(values):
        raise ValueError("Nonempty columns and matching values required")
    # A bound JSON document is a SQL string; explicitly parse it before comparison.
    conditions = [
        f"{ident(c)} <=> CAST(%s AS JSON)"
        if c in json_columns
        else f"{ident(c)} <=> %s"
        for c in columns
    ]
    return " AND ".join(conditions), list(values)


@dataclass(frozen=True)
class Relation:
    child: str
    columns: tuple
    parent: str
    parent_columns: tuple
    source: str
    children: bool = False
    when: tuple = ()

    def applies(self, data):
        return not self.when or data.get(self.when[0]) == self.when[1]


@dataclass
class Table:
    columns: tuple
    signature: tuple
    unique: tuple
    engine: str
    generated: frozenset = frozenset()
    triggers: tuple = ()
    json_columns: frozenset = frozenset()


@dataclass
class Record:
    table: str
    values: tuple
    count: int
    reason: set = field(default_factory=set)
    existing: int = 0
    missing: int = 0
    committed: int = 0
    status: str = "PLANNED"
    note: str = ""


def metadata(cur, schema):
    cur.execute(
        """
        SELECT TABLE_NAME, ENGINE FROM information_schema.TABLES
        WHERE TABLE_SCHEMA = %s AND TABLE_TYPE = 'BASE TABLE'
    """,
        [schema],
    )
    engines = dict(cur.fetchall())
    cur.execute(
        """
        SELECT TABLE_NAME, COLUMN_NAME, COLUMN_TYPE, IS_NULLABLE, EXTRA,
               COLLATION_NAME, GENERATION_EXPRESSION
        FROM information_schema.COLUMNS WHERE TABLE_SCHEMA = %s
        ORDER BY TABLE_NAME, ORDINAL_POSITION
    """,
        [schema],
    )
    columns = defaultdict(list)
    for row in cur.fetchall():
        if row[0] in engines:
            ident(row[0])
            ident(row[1])
            columns[row[0]].append(row[1:])
    cur.execute(
        """
        SELECT TABLE_NAME, INDEX_NAME, COLUMN_NAME, SUB_PART
        FROM information_schema.STATISTICS
        WHERE TABLE_SCHEMA = %s AND NON_UNIQUE = 0
        ORDER BY TABLE_NAME, INDEX_NAME, SEQ_IN_INDEX
    """,
        [schema],
    )
    indexes = defaultdict(lambda: defaultdict(list))
    for table, name, column, prefix in cur.fetchall():
        indexes[table][name].append((column, prefix))
    tables = {}
    for name, cols in columns.items():
        keys = []
        for index, parts in sorted(
            indexes[name].items(), key=lambda x: x[0] != "PRIMARY"
        ):
            if any(c is None or prefix is not None for c, prefix in parts):
                raise ValueError(
                    f"Unsupported expression/prefix unique index: {schema}.{name}.{index}"
                )
            keys.append(tuple(c for c, _ in parts))
        generated = frozenset(c[0] for c in cols if c[5])
        tables[name] = Table(
            tuple(c[0] for c in cols),
            tuple(cols),
            tuple(keys),
            engines[name],
            generated,
            json_columns=frozenset(c[0] for c in cols if c[1].lower() == "json"),
        )
    cur.execute(
        """
        SELECT k.TABLE_NAME, k.CONSTRAINT_NAME, k.COLUMN_NAME,
               k.REFERENCED_TABLE_SCHEMA, k.REFERENCED_TABLE_NAME,
               k.REFERENCED_COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE k
        JOIN information_schema.TABLES t
          ON t.TABLE_SCHEMA = k.TABLE_SCHEMA AND t.TABLE_NAME = k.TABLE_NAME
        WHERE k.TABLE_SCHEMA = %s AND k.REFERENCED_TABLE_NAME IS NOT NULL
          AND t.TABLE_TYPE = 'BASE TABLE'
        ORDER BY k.TABLE_NAME, k.CONSTRAINT_NAME, k.ORDINAL_POSITION
    """,
        [schema],
    )
    groups = defaultdict(list)
    for child, name, col, parent_schema, parent, parent_col in cur.fetchall():
        groups[child, name, parent_schema, parent].append((col, parent_col))
    relations = []
    for (child, name, parent_schema, parent), pairs in groups.items():
        if parent_schema != schema:
            raise ValueError(f"Cross-schema FK unsupported: {schema}.{child}.{name}")
        relations.append(
            Relation(
                child,
                tuple(p[0] for p in pairs),
                parent,
                tuple(p[1] for p in pairs),
                f"FK:{name}",
            )
        )
    cur.execute(
        """
        SELECT EVENT_OBJECT_TABLE, TRIGGER_NAME FROM information_schema.TRIGGERS
        WHERE TRIGGER_SCHEMA = %s
    """,
        [schema],
    )
    triggers = defaultdict(list)
    for table, trigger in cur.fetchall():
        triggers[table].append(trigger)
    for table, names in triggers.items():
        tables[table] = replace(tables[table], triggers=tuple(sorted(names)))
    return tables, relations


def add_legacy(tables, relations):
    result = list(relations)
    specs = list(LEGACY_RELATIONS)
    for table in sorted(LEGACY_COLUMN_TABLES & tables.keys()):
        for column, parent in LEGACY_COLUMNS.items():
            if column in tables[table].columns:
                specs.append((table, (column,), parent, ("id",), False))
    for child, columns, parent, parent_columns, children in specs:
        if child not in tables or not set(columns) <= set(tables[child].columns):
            continue
        # A declared FK on these columns takes precedence over the legacy mapping.
        declared = [r for r in result if r.child == child and r.columns == columns]
        if declared:
            if children:
                for rel in declared:
                    if rel.parent != parent or rel.parent_columns != parent_columns:
                        raise ValueError(
                            f"Reviewed relationship disagrees with FK: {child}{columns}"
                        )
                    result[result.index(rel)] = replace(rel, children=True)
            continue
        if parent not in tables or not set(parent_columns) <= set(
            tables[parent].columns
        ):
            raise ValueError(
                f"Legacy parent missing: {child}{columns} -> {parent}{parent_columns}"
            )
        result.append(
            Relation(child, columns, parent, parent_columns, "LEGACY", children)
        )
    if "test_senario_list_parameter" in tables:
        for column, parent, discriminator, value in (
            ("parent_id", "se_parameter", "category", "SE"),
            ("rflid", "r_parameter", "rflcategory", "R"),
            ("rflid", "f_parameter", "rflcategory", "F"),
            ("rflid", "l_parameter", "rflcategory", "L"),
        ):
            if (
                column in tables["test_senario_list_parameter"].columns
                and parent in tables
            ):
                result.append(
                    Relation(
                        "test_senario_list_parameter",
                        (column,),
                        parent,
                        ("id",),
                        "LEGACY:DISCRIMINATOR",
                        when=(discriminator, value),
                    )
                )
    return result


class CopyPlan:
    def __init__(self, cur, source, target, project, phase):
        self.cur, self.source, self.target = cur, source, target
        self.project, self.phase = project, phase
        self.tables, self.source_fk = metadata(cur, source)
        self.dest, self.dest_fk = metadata(cur, target)
        self.relations = add_legacy(self.tables, self.source_fk)
        self.records = {}
        self.owned = set()
        self.errors = []
        self.coverage = []
        self.project_code = ""
        self.phase_name = ""

    def select(self, schema, table, columns, values):
        where, params = predicate(columns, values, self.tables[table].json_columns)
        cols = ", ".join(ident(c) for c in self.tables[table].columns)
        self.cur.execute(
            f"SELECT {cols} FROM {ident(schema)}.{ident(table)} WHERE {where}", params
        )
        return self.cur.fetchall()

    def add(self, table, rows, reason, owned=False):
        if owned:
            self.owned.add(table)
        counts = Counter(tuple(row) for row in rows)
        changed = False
        for row, count in counts.items():
            key = (table, row)
            if key not in self.records:
                self.records[key] = Record(table, row, count, {reason})
                changed = True
            else:
                record = self.records[key]
                if count > record.count:
                    record.count = count
                    changed = True
                record.reason.add(reason)
        return changed

    def data(self, record):
        return dict(zip(self.tables[record.table].columns, record.values))

    def scope(self, table):
        cols = self.tables[table].columns
        declared = {
            r.columns[0]
            for r in self.source_fk
            if r.child == table
            and r.parent == "project_info"
            and r.parent_columns == ("id",)
            and len(r.columns) == 1
        }
        return [c for c in cols if c in {"project_id", "project_info_id"} | declared]

    def in_scope(self, table, row):
        data = dict(zip(self.tables[table].columns, row))
        if table == "project_info":
            return data["id"] == self.project
        if table == "phase":
            return data["id"] == self.phase
        project_cols = self.scope(table)
        if project_cols and not any(data[c] == self.project for c in project_cols):
            return False
        return "phase_id" not in data or data["phase_id"] == self.phase

    def build(self):
        for table, value, label in (
            ("project_info", self.project, "project_code"),
            ("phase", self.phase, "phase"),
        ):
            if table not in self.tables:
                raise ValueError(f"Missing source table: {table}")
            rows = self.select(self.source, table, ("id",), (value,))
            if len(rows) != 1:
                raise ValueError(f"Source {table}.id={value} must exist exactly once")
            data = dict(zip(self.tables[table].columns, rows[0]))
            if table == "project_info":
                self.project_code = data.get(label, "")
            else:
                self.phase_name = data.get(label, "")
            self.add(table, rows, "INPUT")
        found_scope = False
        for table in sorted(self.tables):
            columns = self.scope(table)
            if not columns:
                continue
            reviewed = table in LEGACY_PROJECT_TABLES or all(
                any(
                    r.child == table
                    and r.columns == (c,)
                    and r.parent == "project_info"
                    for r in self.source_fk
                )
                for c in columns
            )
            # Multiset union preserves duplicate history rows but counts overlap once.
            matching = Counter()
            for column in columns:
                keys, values = [column], [self.project]
                if "phase_id" in self.tables[table].columns:
                    keys.append("phase_id")
                    values.append(self.phase)
                matching |= Counter(
                    tuple(r) for r in self.select(self.source, table, keys, values)
                )
            rows = list(matching.elements())
            self.coverage.append(
                (
                    table,
                    ", ".join(columns),
                    "project + phase"
                    if "phase_id" in self.tables[table].columns
                    else "project-wide",
                    len(rows),
                    "REVIEWED" if reviewed else "UNREVIEWED",
                )
            )
            if rows and not reviewed:
                self.errors.append(
                    f"Unreviewed project relationship: {table}.{columns}"
                )
                continue
            if rows:
                self.add(table, rows, "PROJECT_SCOPE", owned=True)
                if "phase_id" in self.tables[table].columns:
                    found_scope = True
        if not found_scope:
            self.errors.append("No source rows found for the exact project/phase pair")
        queried = set()
        changed = True
        while changed:
            changed = False
            for record in list(self.records.values()):
                data = self.data(record)
                for rel in self.relations:
                    if rel.child == record.table and rel.applies(data):
                        values = tuple(data[c] for c in rel.columns)
                        query = ("P", rel, values)
                        if None not in values and query not in queried:
                            queried.add(query)
                            rows = self.select(
                                self.source, rel.parent, rel.parent_columns, values
                            )
                            if not rows:
                                self.errors.append(
                                    f"Missing source parent: {rel.child}{rel.columns} -> {rel.parent}"
                                )
                            for row in rows:
                                if not self.in_scope(rel.parent, row):
                                    self.errors.append(
                                        f"Dependency crosses selected scope: {rel.child} -> {rel.parent}"
                                    )
                            changed |= self.add(
                                rel.parent, rows, rel.source + ":PARENT"
                            )
                    if rel.parent == record.table and (
                        rel.children
                        or (
                            record.table in self.owned
                            and record.table != "project_info"
                        )
                    ):
                        values = tuple(data[c] for c in rel.parent_columns)
                        query = ("C", rel, values)
                        if None not in values and query not in queried:
                            queried.add(query)
                            rows = self.select(
                                self.source, rel.child, rel.columns, values
                            )
                            rows = [
                                row
                                for row in rows
                                if self.in_scope(rel.child, row)
                                and rel.applies(
                                    dict(zip(self.tables[rel.child].columns, row))
                                )
                            ]
                            changed |= self.add(
                                rel.child,
                                rows,
                                rel.source + ":CHILD",
                                owned=record.table in self.owned,
                            )
                # Reviewed map-name links are strings, not declared foreign keys.
                if record.table in {
                    "project_parameter",
                    "detail_data",
                    "test_project_senario_parameter",
                }:
                    for col in ("value", "updated_value"):
                        name = data.get(col)
                        if not isinstance(name, str) or not name:
                            continue
                        base = re.sub(
                            r"_(?:Ver\.)?[0-9]{3}$", "", name.replace("Ver.", "")
                        )
                        names = (
                            {name, base}
                            | {f"{base}_{n:03d}" for n in range(1, 11)}
                            | {f"{base}_Ver.{n:03d}" for n in range(1, 11)}
                        )
                        for table in (
                            "map_structure",
                            "detail_data_map",
                            "detail_data_map_name",
                        ):
                            if (
                                table not in self.tables
                                or "map_name" not in self.tables[table].columns
                            ):
                                continue
                            for map_name in names:
                                query = ("MAP", table, map_name)
                                if query not in queried:
                                    queried.add(query)
                                    changed |= self.add(
                                        table,
                                        self.select(
                                            self.source,
                                            table,
                                            ("map_name",),
                                            (map_name,),
                                        ),
                                        "LEGACY:MAP_NAME",
                                    )
                if (
                    record.table == "project_r_parameter"
                    and "setup_r_relation" in self.tables
                ):
                    project_row = next(
                        r for r in self.records.values() if r.table == "project_info"
                    )
                    setup = self.data(project_row).get("setup_id")
                    query = ("SETUP_R", setup, data.get("r_parameter_id"))
                    if setup is not None and query not in queried:
                        queried.add(query)
                        changed |= self.add(
                            "setup_r_relation",
                            self.select(
                                self.source,
                                "setup_r_relation",
                                ("setup_id", "r_parameter_id"),
                                query[1:],
                            ),
                            "LEGACY:SETUP_AND_R",
                        )
        # Unknown populated ID columns must be reviewed before claiming completeness.
        for record in self.records.values():
            data = self.data(record)
            covered = {
                c
                for r in self.relations
                if r.child == record.table and r.applies(data)
                for c in r.columns
            }
            for column, value in data.items():
                if value is None or value == "" or column in covered:
                    continue
                if (
                    column.endswith("_id")
                    and column not in {"study_id", "auto_judge_id"}
                    and (column,) not in self.tables[record.table].unique
                ) or column == "rflid":
                    self.errors.append(
                        f"Unmapped populated reference: {record.table}.{column}"
                    )
        self.validate()
        self.errors = list(dict.fromkeys(self.errors))
        if not self.errors:
            try:
                self.insertion_order()
            except ValueError as exc:
                self.errors.append(str(exc))

    def validate(self):
        selected = {r.table for r in self.records.values()}
        for table in selected:
            source = self.tables[table]
            dest = self.dest.get(table)
            if (
                dest is None
                or source.signature != dest.signature
                or source.unique != dest.unique
            ):
                self.errors.append(f"Schema/unique-key mismatch: {table}")
                continue
            if source.engine != "InnoDB" or dest.engine != "InnoDB":
                self.errors.append(f"Transactional InnoDB required: {table}")
            if dest.triggers:
                self.errors.append(
                    f"Target triggers require side-effect review: {table}"
                )
            if source.generated:
                self.errors.append(f"Generated columns need explicit review: {table}")
        source_rel = {
            (r.child, r.columns, r.parent, r.parent_columns)
            for r in self.source_fk
            if r.child in selected
        }
        dest_rel = {
            (r.child, r.columns, r.parent, r.parent_columns)
            for r in self.dest_fk
            if r.child in selected
        }
        if source_rel != dest_rel:
            self.errors.append(
                "Source/target FK definitions differ for selected tables"
            )
        for record in self.records.values():
            if (
                record.table not in self.dest
                or self.tables[record.table].signature
                != self.dest[record.table].signature
                or self.tables[record.table].unique != self.dest[record.table].unique
            ):
                record.status = "BLOCKED_SCHEMA"
                continue
            self.inspect_existing(record)

    def inspect_existing(self, record):
        table = self.tables[record.table]
        data = self.data(record)
        matches = set()
        for columns in table.unique:
            values = tuple(data[c] for c in columns)
            if None in values:
                continue
            matches.update(
                tuple(r)
                for r in self.select(self.target, record.table, columns, values)
            )
        if matches and matches != {record.values}:
            record.status = "CONFLICT"
            record.note = "Unique key exists with different data; no overwrite/remap"
            self.errors.append(
                f"Unique-key conflict: {record.table}; see Records sheet"
            )
            return
        rows = self.select(self.target, record.table, table.columns, record.values)
        if any(tuple(row) != record.values for row in rows):
            record.status = "CONFLICT"
            record.note = (
                "Database comparison matched different bytes/case; review collation"
            )
            self.errors.append(f"Collation/value conflict: {record.table}")
            return
        record.existing = min(len(rows), record.count)
        record.missing = max(record.count - len(rows), 0)
        record.status = "SKIP_EXISTING" if not record.missing else "PLANNED_INSERT"

    def insertion_order(self):
        pending = [r for r in self.records.values() if r.missing]
        dependencies = {}
        for index, record in enumerate(pending):
            deps = set()
            data = self.data(record)
            for rel in self.relations:
                if rel.child != record.table or not rel.applies(data):
                    continue
                values = tuple(data[c] for c in rel.columns)
                if None in values or self.select(
                    self.target, rel.parent, rel.parent_columns, values
                ):
                    continue
                parents = {
                    i
                    for i, parent in enumerate(pending)
                    if parent.table == rel.parent
                    and tuple(self.data(parent)[c] for c in rel.parent_columns)
                    == values
                }
                if not parents:
                    raise ValueError(
                        f"Unresolved target parent: {rel.child} -> {rel.parent}"
                    )
                deps.update(parents - {index})
            dependencies[index] = deps
        order = dependency_order(dependencies)
        return [pending[i] for i in order]

    def execute(self):
        if self.errors:
            raise ValueError("Cannot execute a blocked plan")
        for record in self.insertion_order():
            table = self.tables[record.table]
            cols = ", ".join(ident(c) for c in table.columns)
            placeholders = ", ".join(["%s"] * len(table.columns))
            for _ in range(record.missing):
                self.cur.execute(
                    f"INSERT INTO {ident(self.target)}.{ident(record.table)} "
                    f"({cols}) VALUES ({placeholders})",
                    record.values,
                )
                if self.cur.rowcount != 1:
                    raise ValueError(f"Unexpected INSERT rowcount: {record.table}")
        # Verify rows, including triggers that may have changed supplied values.
        for record in self.records.values():
            rows = self.select(
                self.target,
                record.table,
                self.tables[record.table].columns,
                record.values,
            )
            matched = sum(tuple(r) == record.values for r in rows)
            if matched < record.count:
                table = self.tables[record.table]
                data = self.data(record)
                detail = f"expected={record.count}, matched={matched}"
                for key in table.unique:
                    values = tuple(data[c] for c in key)
                    if None in values:
                        continue
                    candidates = self.select(self.target, record.table, key, values)
                    # Record identifiers and column names only; omit SQL/payload values.
                    detail += (
                        f"; lookup_columns={','.join(key)}, found={len(candidates)}"
                    )
                    if candidates:
                        different = [
                            c
                            for c, actual, expected in zip(
                                table.columns, candidates[0], record.values
                            )
                            if actual != expected
                        ]
                        detail += "; differing_columns=" + (
                            ",".join(different) or "(none)"
                        )
                    break
                record.status = "VERIFICATION_FAILED"
                record.note = detail
                raise ValueError(
                    f"Post-insert verification failed: {record.table}; {detail}"
                )


def dependency_order(dependencies):
    remaining = {key: set(values) for key, values in dependencies.items()}
    order = []
    while remaining:
        ready = [key for key, values in remaining.items() if not values]
        if not ready:
            raise ValueError("Cyclic dependencies require explicit review")
        order.extend(ready)
        for key in ready:
            del remaining[key]
        for values in remaining.values():
            values.difference_update(ready)
    return order


def excel_value(value):
    # Always store strings as text, preventing formula injection from DB contents.
    if value is None or isinstance(value, (int, float, bool)):
        return value
    return str(value)[:32000]


def write_report(path, info, plan):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    wb.remove(wb.active)

    def sheet(name, headers, rows):
        ws = wb.create_sheet(name)
        ws.append(headers)
        for row in rows:
            if ws.max_row >= 1048576:
                raise ValueError("Excel row limit exceeded; reduce input scope")
            ws.append([excel_value(v) for v in row])
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = "s"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="24476B")
        for i, col in enumerate(ws.columns, 1):
            width = min(75, max(14, max(len(str(c.value or "")) for c in col) + 2))
            ws.column_dimensions[get_column_letter(i)].width = width

    records = list(plan.records.values()) if plan else []
    info.update(
        planned_inserted_rows=sum(r.missing for r in records),
        committed_inserted_rows=sum(r.committed for r in records),
        existing_rows=sum(r.existing for r in records),
        committed_deleted_rows=0,
    )
    sheet("Summary", ["Field", "Value"], info.items())
    if plan:
        sheet(
            "Targets",
            [
                "project_id",
                "phase_id",
                "project_code",
                "phase",
                "parent_decision",
                "status",
                "blocking_reasons",
            ],
            [
                (
                    plan.project,
                    plan.phase,
                    plan.project_code,
                    plan.phase_name,
                    "Keep IDs; insert missing parents; skip identical existing rows",
                    info["status"],
                    "\n".join(dict.fromkeys(plan.errors)),
                )
            ],
        )
        by_table = defaultdict(list)
        for record in records:
            by_table[record.table].append(record)
        sheet(
            "Tables",
            [
                "table",
                "relation_sources",
                "source_rows",
                "existing_rows",
                "planned_inserts",
                "committed_inserts",
                "status",
            ],
            [
                (
                    table,
                    ", ".join(sorted(set().union(*(r.reason for r in rows)))),
                    sum(r.count for r in rows),
                    sum(r.existing for r in rows),
                    sum(r.missing for r in rows),
                    sum(r.committed for r in rows),
                    ", ".join(sorted({r.status for r in rows})),
                )
                for table, rows in sorted(by_table.items())
            ],
        )
        sheet(
            "Records",
            [
                "table",
                "key_columns",
                "key_values",
                "source_count",
                "existing_count",
                "planned_inserts",
                "committed_inserts",
                "decision",
                "reason",
                "note",
            ],
            [
                (
                    r.table,
                    ", ".join(plan.tables[r.table].unique[0])
                    if plan.tables[r.table].unique
                    else "(no unique key)",
                    repr(tuple(plan.data(r)[c] for c in plan.tables[r.table].unique[0]))
                    if plan.tables[r.table].unique
                    else "Full-row multiset comparison (payload omitted)",
                    r.count,
                    r.existing,
                    r.missing,
                    r.committed,
                    r.status,
                    ", ".join(sorted(r.reason)),
                    r.note,
                )
                for r in records
            ],
        )
        sheet(
            "Coverage",
            [
                "table",
                "project_columns",
                "scope",
                "matching_rows",
                "relationship_review",
            ],
            plan.coverage,
        )
        sheet(
            "Relations",
            ["child", "columns", "parent", "parent_columns", "source"],
            [
                (
                    r.child,
                    ", ".join(r.columns),
                    r.parent,
                    ", ".join(r.parent_columns),
                    r.source,
                )
                for r in plan.relations
                if r.child in by_table or r.parent in by_table
            ],
        )
        sheet("Issues", ["issue"], [(e,) for e in dict.fromkeys(plan.errors)])
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp.xlsx")
    try:
        wb.save(temporary)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    print(f"Excel report: {path.resolve()}")


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--project-id", type=positive_id, default=RUN_PROJECT_ID)
    parser.add_argument("--phase-id", type=positive_id, default=RUN_PHASE_ID)
    parser.add_argument("--source-db", default=RUN_SOURCE_DATABASE)
    parser.add_argument("--target-db", default=RUN_TARGET_DATABASE)
    parser.add_argument("--execute", action="store_true", default=RUN_EXECUTE)
    parser.add_argument("--report-dir", type=Path, default=RUN_REPORT_DIR)
    args = parser.parse_args(argv)
    if args.project_id is None or args.phase_id is None:
        parser.error(
            "Set RUN_PROJECT_ID/RUN_PHASE_ID or pass --project-id and --phase-id"
        )
    project, phase = positive_id(args.project_id), positive_id(args.phase_id)
    ident(args.source_db)
    ident(args.target_db)
    if args.source_db == args.target_db:
        parser.error("Source and target databases must differ")
    # Validate arguments before Django initialization: --help never opens a DB.
    root = Path(__file__).resolve().parents[2]
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    import django
    from django.apps import apps

    if not apps.ready:
        django.setup()
    from django.db import connection, transaction

    now = datetime.now().astimezone()
    report = (
        args.report_dir.resolve()
        / f"copy_project_{project}_phase_{phase}_{now:%Y%m%d_%H%M%S_%f}.xlsx"
    )
    info = {
        "timestamp": now.isoformat(),
        "mode": "EXECUTE" if args.execute else "DRY_RUN",
        "status": "FAILED",
        "source_database": args.source_db,
        "database": args.target_db,
        "host": connection.settings_dict.get("HOST", ""),
        "input_filename": "(CLI/PyCharm constants)",
        "target_count": 1,
        "project_id": project,
        "phase_id": phase,
        "external_resources": "Database rows only; file/S3 objects are not copied",
        "error": "",
    }
    # Check the report dependency/path before any writes to the database.
    import openpyxl  # noqa: F401

    report.parent.mkdir(parents=True, exist_ok=True)
    plan = None
    exit_code = 0
    try:
        if connection.in_atomic_block or not connection.get_autocommit():
            raise ValueError(
                "Run outside an existing transaction to report commit status correctly"
            )
        with connection.cursor() as cur:
            cur.execute("SET TRANSACTION ISOLATION LEVEL REPEATABLE READ")
        # One repeatable-read snapshot for preview and all writes; no DDL.
        with transaction.atomic():
            with connection.cursor() as cur:
                cur.execute(
                    "SELECT DATABASE(), @@SESSION.foreign_key_checks, @@SESSION.sql_mode"
                )
                database, fk_checks, sql_mode = cur.fetchone()
                if database != args.target_db:
                    raise ValueError(
                        f"Django default database must be {args.target_db!r}; got {database!r}"
                    )
                if not fk_checks or not (
                    {"STRICT_TRANS_TABLES", "STRICT_ALL_TABLES"}
                    & set(sql_mode.split(","))
                ):
                    raise ValueError(
                        "Foreign key checks and strict SQL mode must be enabled"
                    )
                plan = CopyPlan(cur, args.source_db, args.target_db, project, phase)
                plan.build()
                print(
                    f"{args.source_db} -> {database} @ {info['host']}; project={project}, phase={phase}"
                )
                counts = Counter()
                for record in plan.records.values():
                    counts[record.table] += record.missing
                for table, count in sorted(counts.items()):
                    print(f"  {table}: planned inserts={count}")
                if plan.errors:
                    info["status"] = "BLOCKED"
                    raise ValueError("; ".join(dict.fromkeys(plan.errors)))
                if not args.execute:
                    info["status"] = "DRY_RUN"
                elif not any(r.missing for r in plan.records.values()):
                    info["status"] = "NO_OP"
                else:
                    info["status"] = "PREVIEW_READY"
                    write_report(report, info.copy(), plan)
                    token = f"COPY {args.source_db} TO {database} PROJECT {project} PHASE {phase}"
                    print(f"Confirmation required: {token}")
                    try:
                        answer = input("> ")
                    except EOFError:
                        answer = ""
                    if answer != token:
                        info["status"] = "CANCELLED"
                    else:
                        info["status"] = "ROLLING_BACK"
                        plan.execute()
                        info["status"] = "COMMIT_PENDING"
        if info["status"] == "COMMIT_PENDING":
            info["status"] = "COMMITTED"
            for record in plan.records.values():
                record.committed = record.missing
                if record.missing:
                    record.status = "INSERTED"
    except (Exception, KeyboardInterrupt) as exc:
        exit_code = 1
        if info["status"] == "ROLLING_BACK":
            info["status"] = "ROLLED_BACK"
        elif info["status"] == "COMMIT_PENDING":
            info["status"] = "COMMIT_FAILED_OR_UNKNOWN"
        elif info["status"] == "COMMITTED":
            info["status"] = "COMMITTED_REPORTING_INTERRUPTED"
        info["error"] = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        if isinstance(exc, ValueError):
            print(str(exc))
        # Do not print raw driver exceptions/SQL payloads; these can contain private data.
        print(
            f"Run failed ({type(exc).__name__}); status={info['status']}. Review the report."
        )
        if plan:
            plan.errors.append(
                str(exc)
                if isinstance(exc, ValueError)
                else f"{type(exc).__name__}: execution failed; raw driver payload omitted"
            )
    finally:
        if plan and plan.records:
            # Avoid writing raw SQL error messages (which may include payloads) to Excel.
            if exit_code:
                info["error"] = (
                    "Run failed; inspect Issues/Records. Driver payload omitted."
                )
            write_report(report, info, plan)
    print(f"Status: {info['status']}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
