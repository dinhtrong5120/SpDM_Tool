"""Xem trước hoặc xóa dữ liệu của từng cặp project-phase an toàn.

Để chạy bằng PyCharm, sửa RUN_TARGETS/RUN_EXECUTE rồi bấm Run.
Mặc định RUN_EXECUTE=False nên chương trình chỉ xem trước, không xóa dữ liệu.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Iterable, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]

# Khi chạy trực tiếp, phải khởi tạo Django trước khi import database connection.
# Cách này giống utils/update_member_list.py và các script khác trong project.
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

# Cấu hình khi chạy trực tiếp bằng nút Run của PyCharm.
# Mỗi phần tử là một cặp (project_id, phase_id). Không dùng hai danh sách riêng
# vì chúng có thể vô tình tạo tích Descartes và xóa sai project-phase.
# Quy trình an toàn: chạy False trước, kiểm tra kết quả, rồi mới đổi thành True.
RUN_TARGETS = [(1039, 16)]
RUN_EXECUTE = False
RUN_ASSUME_YES = False  # Nên giữ False để xác nhận trong Run console.

# Các bảng legacy có project_id nhưng chưa khai báo FK tới project_info.id.
# Chỉ bổ sung bảng sau khi đã kiểm tra ý nghĩa của cột trong schema.
LEGACY_PROJECT_COLUMNS: dict[str, tuple[str, ...]] = {
    "detail_data": ("project_id",),
    "detail_data_rlist": ("project_id",),
    "project_parameter_attachments": ("project_id",),
    "project_r_parameter_manager_approval_history": ("project_id",),
    "studies_comparison_dashboard_settings": ("project_id",),
    "temp_project_parameter": ("project_id",),
}

IDENTIFIER_PATTERN = re.compile(r"^[A-Za-z0-9_$]+$")


def quote_identifier(identifier: str) -> str:
    """Kiểm tra rồi quote tên bảng/cột lấy từ metadata."""
    if not IDENTIFIER_PATTERN.fullmatch(identifier):
        raise ValueError(f"Tên bảng/cột không hợp lệ: {identifier!r}")
    return f"`{identifier}`"


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Không thể tạo SQL placeholders cho danh sách rỗng.")
    return ", ".join(["%s"] * len(values))


def normalize_positive_id(raw_id: int, *, label: str) -> int:
    if isinstance(raw_id, bool):
        raise ValueError(f"{label} phải là số nguyên dương, không phải boolean.")
    try:
        normalized_id = int(raw_id)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} không hợp lệ: {raw_id!r}") from exc
    if normalized_id <= 0:
        raise ValueError(f"{label} phải lớn hơn 0: {normalized_id}")
    return normalized_id


def normalize_targets(
    targets: Iterable[Sequence[int]],
) -> list[tuple[int, int]]:
    """Chuẩn hóa các cặp (project_id, phase_id), không trùng lặp."""
    normalized: list[tuple[int, int]] = []
    seen: set[tuple[int, int]] = set()

    for target in targets:
        try:
            project_raw, phase_raw = target
        except (TypeError, ValueError) as exc:
            raise ValueError(
                "Mỗi target phải có đúng dạng (project_id, phase_id)."
            ) from exc
        pair = (
            normalize_positive_id(project_raw, label="project_id"),
            normalize_positive_id(phase_raw, label="phase_id"),
        )
        if pair not in seen:
            normalized.append(pair)
            seen.add(pair)

    if not normalized:
        raise ValueError("Danh sách project-phase targets rỗng.")
    return normalized


def get_existing_legacy_columns(cursor) -> dict[str, list[str]]:
    """Lấy cột allowlist ở bảng có phase_id trong database hiện tại."""
    if not LEGACY_PROJECT_COLUMNS:
        return {}

    table_names = sorted(LEGACY_PROJECT_COLUMNS)
    placeholders = make_placeholders(table_names)
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
        table_names,
    )

    existing: dict[str, list[str]] = defaultdict(list)
    for table_name, column_name in cursor.fetchall():
        if column_name in LEGACY_PROJECT_COLUMNS[table_name]:
            existing[table_name].append(column_name)
    return dict(existing)


def get_phase_related_tables(cursor) -> list[dict[str, object]]:
    """Lấy bảng có cả cột project liên quan và phase_id."""
    cursor.execute(
        """
        SELECT DISTINCT
            kcu.TABLE_NAME,
            kcu.COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE AS kcu
        JOIN information_schema.COLUMNS AS phase_column
          ON phase_column.TABLE_SCHEMA = kcu.TABLE_SCHEMA
         AND phase_column.TABLE_NAME = kcu.TABLE_NAME
         AND phase_column.COLUMN_NAME = 'phase_id'
        JOIN information_schema.TABLES AS t
          ON t.TABLE_SCHEMA = kcu.TABLE_SCHEMA
         AND t.TABLE_NAME = kcu.TABLE_NAME
        WHERE kcu.TABLE_SCHEMA = DATABASE()
          AND t.TABLE_TYPE = 'BASE TABLE'
          AND kcu.TABLE_NAME <> %s
          AND kcu.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND kcu.REFERENCED_TABLE_NAME = %s
          AND kcu.REFERENCED_COLUMN_NAME = %s
        ORDER BY kcu.TABLE_NAME, kcu.COLUMN_NAME
        """,
        (PROJECT_TABLE, PROJECT_TABLE, PROJECT_KEY),
    )

    table_columns: dict[str, set[str]] = defaultdict(set)
    relation_types: dict[str, set[str]] = defaultdict(set)
    for table_name, column_name in cursor.fetchall():
        table_columns[table_name].add(column_name)
        relation_types[table_name].add("foreign_key")

    for table_name, columns in get_existing_legacy_columns(cursor).items():
        table_columns[table_name].update(columns)
        relation_types[table_name].add("legacy_allowlist")

    return [
        {
            "table_name": table_name,
            "project_columns": sorted(columns),
            "phase_column": "phase_id",
            "relation_types": sorted(relation_types[table_name]),
        }
        for table_name, columns in sorted(table_columns.items())
    ]


def order_tables_child_first(
    cursor,
    related_tables: Sequence[dict[str, object]],
) -> list[dict[str, object]]:
    """Sắp xếp để bảng con được xóa trước bảng cha."""
    by_name = {str(item["table_name"]): item for item in related_tables}
    if not by_name:
        return []

    table_names = sorted(by_name)
    placeholders = make_placeholders(table_names)
    params = [*table_names, *table_names]
    cursor.execute(
        f"""
        SELECT DISTINCT TABLE_NAME, REFERENCED_TABLE_NAME
        FROM information_schema.KEY_COLUMN_USAGE
        WHERE TABLE_SCHEMA = DATABASE()
          AND REFERENCED_TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME IN ({placeholders})
          AND REFERENCED_TABLE_NAME IN ({placeholders})
        """,
        params,
    )

    children: dict[str, set[str]] = defaultdict(set)
    for child_table, parent_table in cursor.fetchall():
        if child_table != parent_table:
            children[parent_table].add(child_table)

    ordered_names: list[str] = []
    visiting: set[str] = set()
    visited: set[str] = set()

    def visit(table_name: str) -> None:
        if table_name in visited:
            return
        if table_name in visiting:
            raise RuntimeError(
                f"Phát hiện vòng lặp FK giữa các bảng liên quan tại {table_name}."
            )
        visiting.add(table_name)
        for child in sorted(children.get(table_name, ())):
            visit(child)
        visiting.remove(table_name)
        visited.add(table_name)
        ordered_names.append(table_name)

    for table_name in table_names:
        visit(table_name)
    return [by_name[name] for name in ordered_names]


def build_target_filter(
    project_columns: Sequence[str],
    phase_column: str,
    targets: Sequence[tuple[int, int]],
) -> tuple[str, tuple[int, ...]]:
    """Tạo điều kiện chỉ khớp chính xác từng cặp project-phase."""
    if not project_columns:
        raise ValueError("Không có cột project để tạo điều kiện SQL.")
    if not targets:
        raise ValueError("Không có project-phase target để tạo điều kiện SQL.")

    phase_identifier = quote_identifier(phase_column)
    target_clauses: list[str] = []
    params: list[int] = []
    for project_id, phase_id in targets:
        project_clauses = [
            f"{quote_identifier(column)} = %s" for column in project_columns
        ]
        target_clauses.append(
            f"(({' OR '.join(project_clauses)}) AND {phase_identifier} = %s)"
        )
        params.extend([project_id] * len(project_columns))
        params.append(phase_id)
    return " OR ".join(target_clauses), tuple(params)


def find_projects(cursor, project_ids: Sequence[int]) -> list[dict[str, object]]:
    placeholders = make_placeholders(project_ids)
    cursor.execute(
        f"""
        SELECT id, project_code
        FROM {quote_identifier(PROJECT_TABLE)}
        WHERE {quote_identifier(PROJECT_KEY)} IN ({placeholders})
        ORDER BY id
        """,
        project_ids,
    )
    return [
        {"id": project_id, "project_code": project_code}
        for project_id, project_code in cursor.fetchall()
    ]


def find_phases(cursor, phase_ids: Sequence[int]) -> list[dict[str, object]]:
    placeholders = make_placeholders(phase_ids)
    cursor.execute(
        f"SELECT id, phase FROM phase WHERE id IN ({placeholders}) ORDER BY id",
        phase_ids,
    )
    return [
        {"id": phase_id, "phase": phase_name}
        for phase_id, phase_name in cursor.fetchall()
    ]


def count_related_rows(
    cursor,
    related_tables: Sequence[dict[str, object]],
    targets: Sequence[tuple[int, int]],
) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in related_tables:
        table_name = str(item["table_name"])
        project_columns = list(item["project_columns"])
        phase_column = str(item["phase_column"])
        where_sql, params = build_target_filter(
            project_columns, phase_column, targets
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
    projects: Sequence[dict[str, object]],
    phases: Sequence[dict[str, object]],
    targets: Sequence[tuple[int, int]],
    related_tables: Sequence[dict[str, object]],
    counts: dict[str, int],
) -> None:
    settings = db_connection.settings_dict
    print(f"Database: {settings.get('NAME')}")
    print(f"Host: {settings.get('HOST')}:{settings.get('PORT') or 3306}")
    print("Projects tìm thấy:")
    for project in projects:
        print(f"- {project['id']}: {project['project_code']}")

    print("Phases tìm thấy:")
    for phase in phases:
        print(f"- {phase['id']}: {phase['phase']}")

    print("Targets sẽ xóa:")
    for project_id, phase_id in targets:
        print(f"- project_id={project_id}, phase_id={phase_id}")

    print("\nCác bảng sẽ được xử lý (con trước, cha sau):")
    for item in related_tables:
        table_name = str(item["table_name"])
        columns = ", ".join(
            f"`{column}`" for column in item["project_columns"]
        )
        relation_types = ", ".join(item["relation_types"])
        print(
            f"- {table_name}: {counts[table_name]} rows; "
            f"project_columns={columns}; phase_column=`{item['phase_column']}`; "
            f"source={relation_types}"
        )
    print(f"Tổng số dòng liên quan: {sum(counts.values())}")


def delete_project_phase_data(
    targets: Iterable[Sequence[int]],
    db_connection=connection,
    *,
    execute: bool = False,
    assume_yes: bool = False,
) -> dict[str, int]:
    """Xóa dữ liệu đúng cặp project-phase; không xóa project_info."""
    normalized_targets = normalize_targets(targets)
    project_ids = list(dict.fromkeys(item[0] for item in normalized_targets))
    phase_ids = list(dict.fromkeys(item[1] for item in normalized_targets))

    with db_connection.cursor() as cursor:
        projects = find_projects(cursor, project_ids)
        found_ids = {int(project["id"]) for project in projects}
        missing_ids = [
            project_id
            for project_id in project_ids
            if project_id not in found_ids
        ]
        if missing_ids:
            raise ValueError(f"Không tìm thấy project IDs: {missing_ids}")

        phases = find_phases(cursor, phase_ids)
        found_phase_ids = {int(phase["id"]) for phase in phases}
        missing_phase_ids = [
            phase_id for phase_id in phase_ids if phase_id not in found_phase_ids
        ]
        if missing_phase_ids:
            raise ValueError(f"Không tìm thấy phase IDs: {missing_phase_ids}")

        related_tables = order_tables_child_first(
            cursor, get_phase_related_tables(cursor)
        )
        counts = count_related_rows(cursor, related_tables, normalized_targets)
        print_plan(
            db_connection,
            projects,
            phases,
            normalized_targets,
            related_tables,
            counts,
        )
        if not any(counts.values()):
            raise ValueError(
                "Không tìm thấy dữ liệu nào thuộc các cặp project-phase đã chọn."
            )

    if not execute:
        print("\nDRY RUN: chưa có dữ liệu nào bị xóa. Đặt RUN_EXECUTE=True để xóa.")
        return counts

    database_name = db_connection.settings_dict.get("NAME") or "<unknown>"
    target_label = ",".join(
        f"{project_id}:{phase_id}"
        for project_id, phase_id in normalized_targets
    )
    confirmation = f"DELETE_PHASE_DATA {database_name}:{target_label}"
    if not assume_yes:
        answer = input(f"\nNhập chính xác '{confirmation}' để tiếp tục: ").strip()
        if answer != confirmation:
            print("Đã hủy. Không có dữ liệu nào bị xóa.")
            return counts

    deleted_counts: dict[str, int] = {}
    with transaction.atomic(using=db_connection.alias):
        with db_connection.cursor() as cursor:
            # Đọc lại metadata để hạn chế schema thay đổi giữa preview và lúc xóa.
            related_tables = order_tables_child_first(
                cursor, get_phase_related_tables(cursor)
            )
            for item in related_tables:
                table_name = str(item["table_name"])
                project_columns = list(item["project_columns"])
                phase_column = str(item["phase_column"])
                where_sql, params = build_target_filter(
                    project_columns, phase_column, normalized_targets
                )
                cursor.execute(
                    f"SELECT COUNT(*) FROM {quote_identifier(table_name)} "
                    f"WHERE {where_sql} FOR UPDATE",
                    params,
                )
                expected = int(cursor.fetchone()[0])
                cursor.execute(
                    f"DELETE FROM {quote_identifier(table_name)} "
                    f"WHERE {where_sql}",
                    params,
                )
                actual = int(cursor.rowcount)
                if actual != expected:
                    raise RuntimeError(
                        f"Số dòng xóa ở {table_name} không đúng; transaction "
                        f"sẽ rollback: expected={expected}, actual={actual}"
                    )
                deleted_counts[table_name] = actual
                print(f"{table_name}: deleted {actual} rows")

            remaining = count_related_rows(
                cursor, related_tables, normalized_targets
            )
            not_empty = {
                table: count for table, count in remaining.items() if count
            }
            if not_empty:
                raise RuntimeError(
                    "Vẫn còn dữ liệu tham chiếu trực tiếp; transaction sẽ "
                    f"rollback: {not_empty}"
                )

    print("\nXóa dữ liệu project-phase thành công; project_info được giữ nguyên.")
    return deleted_counts


def parse_target_arg(value: str) -> tuple[int, int]:
    try:
        project_text, phase_text = value.split(":", maxsplit=1)
        return (
            normalize_positive_id(project_text, label="project_id"),
            normalize_positive_id(phase_text, label="phase_id"),
        )
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Target phải có dạng PROJECT_ID:PHASE_ID, nhận được {value!r}."
        ) from exc


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Preview hoặc xóa đúng dữ liệu của từng cặp project-phase."
    )
    parser.add_argument(
        "--targets",
        nargs="+",
        type=parse_target_arg,
        metavar="PROJECT_ID:PHASE_ID",
        help="Một hoặc nhiều cặp project-phase; ghi đè RUN_TARGETS.",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Thực hiện DELETE; nếu bỏ qua thì chương trình chỉ dry-run.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Bỏ qua xác nhận tương tác; chỉ dùng cho automation đã kiểm soát.",
    )
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    targets = args.targets if args.targets is not None else RUN_TARGETS
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES

    if assume_yes and not execute:
        raise ValueError("RUN_ASSUME_YES/--yes chỉ hợp lệ khi bật chế độ xóa.")

    delete_project_phase_data(
        targets,
        execute=execute,
        assume_yes=assume_yes,
    )


if __name__ == "__main__":
    main()
