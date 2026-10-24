"""Xem trước hoặc xóa dữ liệu variation theo project và phase an toàn.

Script chỉ xóa các dòng thuộc đúng bộ ``project_id + phase_id + variation_id``
trong bốn bảng được khai báo ở ``TARGET_TABLES``. Bản ghi master trong bảng
``variation`` không bị xóa vì có thể vẫn được project/phase khác sử dụng.

Để chạy bằng nút Run của PyCharm, sửa các hằng ``RUN_*`` bên dưới. Mặc định
``RUN_EXECUTE = False`` nên chương trình chỉ xem trước và xuất báo cáo Excel.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable, Sequence


PROJECT_ROOT = Path(__file__).resolve().parents[2]

# Khởi tạo Django trước khi import connection nếu file được chạy trực tiếp.
if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402


# =========================================================
# CẤU HÌNH CHẠY BẰNG PYCHARM
# =========================================================

RUN_PROJECT_ID = 1020
RUN_PHASE_ID = 35
RUN_VARIATION_IDS = [999996]

# False: chỉ xem trước và xuất Excel. True: cho phép đi tới bước xác nhận xóa.
RUN_EXECUTE = False

# Nên luôn giữ False. Khi True, script bỏ qua xác nhận tương tác chính xác.
RUN_ASSUME_YES = False

# None: tự tạo file có timestamp trong utils/Tools_new/output/.
RUN_OUTPUT_PATH: str | None = None

# Giới hạn số dòng chi tiết ghi vào mỗi sheet; tổng số dòng thật vẫn có ở Summary.
RUN_REPORT_ROW_LIMIT = 50_000


# Chỉ các bảng đã được review cho nghiệp vụ này mới được phép xóa.
TARGET_TABLES = (
    "detail_data",
    "detail_data_rlist",
    "project_parameter",
    "project_r_parameter",
)

IDENTIFIER_PATTERN = re.compile(r"^[A-Za-z0-9_$]+$")
EXCEL_ILLEGAL_CHARACTER_PATTERN = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")
EXCEL_CELL_LIMIT = 32_767


@dataclass(frozen=True)
class RunConfig:
    project_id: int
    phase_id: int
    variation_ids: tuple[int, ...]
    execute: bool
    assume_yes: bool
    output_path: Path
    report_row_limit: int


@dataclass
class TableSnapshot:
    table_name: str
    columns: list[str]
    rows: list[tuple[Any, ...]]
    total_rows: int

    @property
    def truncated(self) -> bool:
        return self.total_rows > len(self.rows)


@dataclass(frozen=True)
class RelatedImpact:
    constraint_name: str
    table_name: str
    parent_table: str
    delete_rule: str
    action: str
    row_count: int | None
    note: str


def quote_identifier(identifier: str) -> str:
    """Kiểm tra và quote tên bảng/cột lấy từ allowlist hoặc metadata."""
    if not IDENTIFIER_PATTERN.fullmatch(identifier):
        raise ValueError(f"Tên bảng/cột không hợp lệ: {identifier!r}")
    return f"`{identifier}`"


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Không thể tạo placeholders cho danh sách rỗng.")
    return ", ".join(["%s"] * len(values))


def normalize_positive_id(raw_id: Any, *, label: str) -> int:
    if isinstance(raw_id, bool):
        raise ValueError(f"{label} phải là số nguyên dương, không phải boolean.")
    try:
        normalized_id = int(raw_id)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} không hợp lệ: {raw_id!r}") from exc
    if normalized_id <= 0:
        raise ValueError(f"{label} phải lớn hơn 0: {normalized_id}")
    return normalized_id


def normalize_variation_ids(raw_ids: Iterable[Any]) -> tuple[int, ...]:
    """Chuẩn hóa ID dương, loại trùng nhưng giữ nguyên thứ tự nhập."""
    normalized: list[int] = []
    seen: set[int] = set()
    for raw_id in raw_ids:
        variation_id = normalize_positive_id(raw_id, label="variation_id")
        if variation_id not in seen:
            normalized.append(variation_id)
            seen.add(variation_id)
    if not normalized:
        raise ValueError("VARIATION_IDS không được để trống.")
    return tuple(normalized)


def normalize_report_row_limit(raw_limit: Any) -> int:
    limit = normalize_positive_id(raw_limit, label="report_row_limit")
    if limit > 1_000_000:
        raise ValueError("report_row_limit không được lớn hơn 1.000.000.")
    return limit


def build_target_filter(
    project_id: int,
    phase_id: int,
    variation_ids: Sequence[int],
) -> tuple[str, list[int]]:
    placeholders = make_placeholders(variation_ids)
    where_sql = (
        f"project_id = %s AND phase_id = %s AND variation_id IN ({placeholders})"
    )
    return where_sql, [project_id, phase_id, *variation_ids]


def default_output_path(project_id: int, phase_id: int) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = (
        f"variation_delete_report_project_{project_id}_phase_{phase_id}_"
        f"{timestamp}.xlsx"
    )
    return Path(__file__).resolve().parent / "output" / filename


def resolve_output_path(
    raw_path: str | None,
    *,
    project_id: int,
    phase_id: int,
) -> Path:
    if raw_path:
        path = Path(raw_path).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        if path.suffix.lower() != ".xlsx":
            raise ValueError("Đường dẫn báo cáo phải có phần mở rộng .xlsx.")
        return path.resolve()
    return default_output_path(project_id, phase_id)


def parse_args(argv: Sequence[str] | None = None) -> RunConfig:
    parser = argparse.ArgumentParser(
        description=(
            "Xem trước/xóa dữ liệu variation theo project và phase, đồng thời "
            "xuất báo cáo Excel."
        )
    )
    parser.add_argument("--project-id", type=int, default=RUN_PROJECT_ID)
    parser.add_argument("--phase-id", type=int, default=RUN_PHASE_ID)
    parser.add_argument(
        "--variation-ids",
        nargs="+",
        type=int,
        default=RUN_VARIATION_IDS,
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        default=RUN_EXECUTE,
        help="Cho phép thực hiện xóa sau bước xác nhận.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        default=RUN_ASSUME_YES,
        help="Bỏ qua xác nhận tương tác (không khuyến nghị).",
    )
    parser.add_argument("--output", default=RUN_OUTPUT_PATH)
    parser.add_argument(
        "--report-row-limit",
        type=int,
        default=RUN_REPORT_ROW_LIMIT,
    )
    args = parser.parse_args(argv)

    project_id = normalize_positive_id(args.project_id, label="project_id")
    phase_id = normalize_positive_id(args.phase_id, label="phase_id")
    variation_ids = normalize_variation_ids(args.variation_ids)
    report_row_limit = normalize_report_row_limit(args.report_row_limit)

    return RunConfig(
        project_id=project_id,
        phase_id=phase_id,
        variation_ids=variation_ids,
        execute=bool(args.execute),
        assume_yes=bool(args.yes),
        output_path=resolve_output_path(
            args.output,
            project_id=project_id,
            phase_id=phase_id,
        ),
        report_row_limit=report_row_limit,
    )


def database_identity() -> tuple[str, str]:
    settings_dict = connection.settings_dict
    return str(settings_dict.get("NAME") or ""), str(settings_dict.get("HOST") or "")


def validate_schema(cursor) -> None:
    """Chỉ tiếp tục khi bốn bảng có đủ ba cột giới hạn phạm vi xóa."""
    placeholders = make_placeholders(TARGET_TABLES)
    cursor.execute(
        f"""
        SELECT c.TABLE_NAME, c.COLUMN_NAME
        FROM information_schema.COLUMNS AS c
        JOIN information_schema.TABLES AS t
          ON t.TABLE_SCHEMA = c.TABLE_SCHEMA
         AND t.TABLE_NAME = c.TABLE_NAME
         AND t.TABLE_TYPE = 'BASE TABLE'
        WHERE c.TABLE_SCHEMA = DATABASE()
          AND c.TABLE_NAME IN ({placeholders})
          AND c.COLUMN_NAME IN ('project_id', 'phase_id', 'variation_id')
        """,
        list(TARGET_TABLES),
    )

    columns_by_table = {table: set() for table in TARGET_TABLES}
    for table_name, column_name in cursor.fetchall():
        columns_by_table[table_name].add(column_name)

    required = {"project_id", "phase_id", "variation_id"}
    errors = []
    for table_name in TARGET_TABLES:
        missing = sorted(required - columns_by_table[table_name])
        if missing:
            errors.append(f"{table_name}: thiếu {', '.join(missing)}")
    if errors:
        raise RuntimeError(
            "Schema database không phù hợp với script: " + "; ".join(errors)
        )


def validate_project_and_phase(cursor, project_id: int, phase_id: int) -> None:
    cursor.execute("SELECT 1 FROM project_info WHERE id = %s", [project_id])
    if cursor.fetchone() is None:
        raise ValueError(f"project_id={project_id} không tồn tại.")

    cursor.execute("SELECT 1 FROM phase WHERE id = %s", [phase_id])
    if cursor.fetchone() is None:
        raise ValueError(f"phase_id={phase_id} không tồn tại.")


def fetch_variations(cursor, variation_ids: Sequence[int]) -> list[dict[str, Any]]:
    placeholders = make_placeholders(variation_ids)
    cursor.execute(
        f"""
        SELECT id, variation
        FROM variation
        WHERE id IN ({placeholders})
        ORDER BY id
        """,
        list(variation_ids),
    )
    labels = {int(row[0]): row[1] for row in cursor.fetchall()}
    return [
        {
            "variation_id": variation_id,
            "variation": labels.get(variation_id),
            "exists_in_master": variation_id in labels,
        }
        for variation_id in variation_ids
    ]


def count_target_rows(
    cursor,
    project_id: int,
    phase_id: int,
    variation_ids: Sequence[int],
) -> dict[str, int]:
    where_sql, params = build_target_filter(project_id, phase_id, variation_ids)
    counts: dict[str, int] = {}
    for table_name in TARGET_TABLES:
        cursor.execute(
            f"SELECT COUNT(*) FROM {quote_identifier(table_name)} WHERE {where_sql}",
            params,
        )
        counts[table_name] = int(cursor.fetchone()[0])
    return counts


def fetch_used_variation_ids(
    cursor,
    project_id: int,
    phase_id: int,
    variation_ids: Sequence[int],
) -> set[int]:
    where_sql, params = build_target_filter(project_id, phase_id, variation_ids)
    union_parts = [
        f"SELECT variation_id FROM {quote_identifier(table)} WHERE {where_sql}"
        for table in TARGET_TABLES
    ]
    cursor.execute(" UNION ".join(union_parts), params * len(TARGET_TABLES))
    return {int(row[0]) for row in cursor.fetchall()}


def fetch_table_snapshot(
    cursor,
    table_name: str,
    project_id: int,
    phase_id: int,
    variation_ids: Sequence[int],
    row_limit: int,
) -> TableSnapshot:
    where_sql, params = build_target_filter(project_id, phase_id, variation_ids)
    cursor.execute(
        f"SELECT * FROM {quote_identifier(table_name)} WHERE {where_sql} LIMIT %s",
        [*params, row_limit],
    )
    columns = [description[0] for description in cursor.description]
    rows = list(cursor.fetchall())

    cursor.execute(
        f"SELECT COUNT(*) FROM {quote_identifier(table_name)} WHERE {where_sql}",
        params,
    )
    total_rows = int(cursor.fetchone()[0])
    return TableSnapshot(table_name, columns, rows, total_rows)


def fetch_fk_dependencies(cursor) -> list[dict[str, Any]]:
    """Liệt kê FK trỏ tới bảng bị xóa và quy tắc ON DELETE để audit."""
    placeholders = make_placeholders(TARGET_TABLES)
    cursor.execute(
        f"""
        SELECT
            kcu.CONSTRAINT_NAME,
            kcu.TABLE_NAME AS child_table,
            GROUP_CONCAT(
                kcu.COLUMN_NAME ORDER BY kcu.ORDINAL_POSITION SEPARATOR ', '
            ) AS child_columns,
            kcu.REFERENCED_TABLE_NAME AS parent_table,
            GROUP_CONCAT(
                kcu.REFERENCED_COLUMN_NAME
                ORDER BY kcu.ORDINAL_POSITION SEPARATOR ', '
            ) AS parent_columns,
            rc.DELETE_RULE
        FROM information_schema.KEY_COLUMN_USAGE AS kcu
        JOIN information_schema.REFERENTIAL_CONSTRAINTS AS rc
          ON rc.CONSTRAINT_SCHEMA = kcu.CONSTRAINT_SCHEMA
         AND rc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME
         AND rc.TABLE_NAME = kcu.TABLE_NAME
        WHERE kcu.TABLE_SCHEMA = DATABASE()
          AND kcu.REFERENCED_TABLE_SCHEMA = DATABASE()
          AND kcu.REFERENCED_TABLE_NAME IN ({placeholders})
        GROUP BY
            kcu.CONSTRAINT_NAME,
            kcu.TABLE_NAME,
            kcu.REFERENCED_TABLE_NAME,
            rc.DELETE_RULE
        ORDER BY parent_table, child_table, kcu.CONSTRAINT_NAME
        """,
        list(TARGET_TABLES),
    )
    return [
        {
            "constraint_name": row[0],
            "child_table": row[1],
            "child_columns": row[2],
            "parent_table": row[3],
            "parent_columns": row[4],
            "delete_rule": row[5],
        }
        for row in cursor.fetchall()
    ]


def fetch_table_columns(cursor, table_names: Sequence[str]) -> dict[str, set[str]]:
    if not table_names:
        return {}
    placeholders = make_placeholders(table_names)
    cursor.execute(
        f"""
        SELECT c.TABLE_NAME, c.COLUMN_NAME
        FROM information_schema.COLUMNS AS c
        JOIN information_schema.TABLES AS t
          ON t.TABLE_SCHEMA = c.TABLE_SCHEMA
         AND t.TABLE_NAME = c.TABLE_NAME
         AND t.TABLE_TYPE = 'BASE TABLE'
        WHERE c.TABLE_SCHEMA = DATABASE()
          AND c.TABLE_NAME IN ({placeholders})
        """,
        list(table_names),
    )
    columns_by_table = {table_name: set() for table_name in table_names}
    for table_name, column_name in cursor.fetchall():
        columns_by_table[table_name].add(column_name)
    return columns_by_table


def fetch_related_impacts(
    cursor,
    dependencies: Sequence[dict[str, Any]],
    project_id: int,
    phase_id: int,
    variation_ids: Sequence[int],
) -> list[RelatedImpact]:
    """Ước lượng các dòng ngoài allowlist bị FK tác động hoặc chặn xóa."""
    related_tables = sorted(
        {
            str(dependency["child_table"])
            for dependency in dependencies
            if dependency["child_table"] not in TARGET_TABLES
        }
    )
    columns_by_table = fetch_table_columns(cursor, related_tables)
    required_columns = {"project_id", "phase_id", "variation_id"}
    where_sql, params = build_target_filter(project_id, phase_id, variation_ids)
    impacts: list[RelatedImpact] = []

    for dependency in dependencies:
        table_name = str(dependency["child_table"])
        if table_name in TARGET_TABLES:
            continue
        delete_rule = str(dependency["delete_rule"]).upper()
        if delete_rule == "CASCADE":
            action = "DELETE (FK CASCADE)"
        elif delete_rule == "SET NULL":
            action = "UPDATE (FK SET NULL)"
        else:
            action = "POTENTIAL BLOCK"

        missing_columns = sorted(
            required_columns - columns_by_table.get(table_name, set())
        )
        if missing_columns:
            impacts.append(
                RelatedImpact(
                    constraint_name=str(dependency["constraint_name"]),
                    table_name=table_name,
                    parent_table=str(dependency["parent_table"]),
                    delete_rule=delete_rule,
                    action=action,
                    row_count=None,
                    note=(
                        "Không thể ước lượng an toàn vì thiếu cột phạm vi: "
                        + ", ".join(missing_columns)
                    ),
                )
            )
            continue

        cursor.execute(
            f"SELECT COUNT(*) FROM {quote_identifier(table_name)} WHERE {where_sql}",
            params,
        )
        row_count = int(cursor.fetchone()[0])
        impacts.append(
            RelatedImpact(
                constraint_name=str(dependency["constraint_name"]),
                table_name=table_name,
                parent_table=str(dependency["parent_table"]),
                delete_rule=delete_rule,
                action=action,
                row_count=row_count,
                note=(
                    "Chỉ tác động đến bản ghi database; script không xóa file/S3."
                    if table_name == "project_parameter_attachments"
                    else ""
                ),
            )
        )
    return impacts


def order_tables_child_first(
    tables: Sequence[str],
    dependencies: Sequence[dict[str, Any]],
) -> list[str]:
    """Sắp xếp bảng con trước bảng cha đối với FK nằm trong allowlist."""
    allowed = set(tables)
    children: dict[str, set[str]] = {table: set() for table in tables}
    for dependency in dependencies:
        child = str(dependency["child_table"])
        parent = str(dependency["parent_table"])
        if child in allowed and parent in allowed and child != parent:
            children[parent].add(child)

    ordered: list[str] = []
    visiting: set[str] = set()
    visited: set[str] = set()

    def visit(table_name: str) -> None:
        if table_name in visited:
            return
        if table_name in visiting:
            raise RuntimeError(f"Phát hiện vòng lặp FK tại {table_name}.")
        visiting.add(table_name)
        for child in sorted(children[table_name]):
            visit(child)
        visiting.remove(table_name)
        visited.add(table_name)
        ordered.append(table_name)

    for table_name in tables:
        visit(table_name)
    return ordered


def delete_target_rows(
    config: RunConfig,
    before_counts: dict[str, int],
    delete_order: Sequence[str],
) -> tuple[dict[str, int], dict[str, int]]:
    where_sql, params = build_target_filter(
        config.project_id,
        config.phase_id,
        config.variation_ids,
    )
    deleted_counts: dict[str, int] = {}

    with transaction.atomic():
        with connection.cursor() as cursor:
            for table_name in delete_order:
                cursor.execute(
                    f"DELETE FROM {quote_identifier(table_name)} WHERE {where_sql}",
                    params,
                )
                deleted_count = int(cursor.rowcount)
                expected_count = before_counts[table_name]
                if deleted_count != expected_count:
                    raise RuntimeError(
                        f"{table_name}: dự kiến xóa {expected_count} dòng nhưng "
                        f"database báo {deleted_count} dòng. Đã rollback."
                    )
                deleted_counts[table_name] = deleted_count

            after_counts = count_target_rows(
                cursor,
                config.project_id,
                config.phase_id,
                config.variation_ids,
            )
            remaining = {table: count for table, count in after_counts.items() if count}
            if remaining:
                raise RuntimeError(
                    f"Vẫn còn dữ liệu sau khi xóa: {remaining}. Đã rollback."
                )

    return deleted_counts, after_counts


def excel_value(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, datetime, date)):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, bytes):
        text_value = value.hex()
    else:
        text_value = str(value)
    text_value = EXCEL_ILLEGAL_CHARACTER_PATTERN.sub("�", text_value)
    if text_value.startswith(("=", "+", "-", "@")):
        text_value = "'" + text_value
    if len(text_value) > EXCEL_CELL_LIMIT:
        return text_value[: EXCEL_CELL_LIMIT - 16] + "...[TRUNCATED]"
    return text_value


def style_sheet(worksheet) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        max_length = max(
            (len(str(cell.value)) for cell in column_cells if cell.value is not None),
            default=0,
        )
        column_letter = get_column_letter(column_cells[0].column)
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 60)


def append_mapping_rows(worksheet, rows: Sequence[dict[str, Any]]) -> None:
    if not rows:
        worksheet.append(["Không có dữ liệu"])
        return
    headers = list(rows[0])
    worksheet.append(headers)
    for row in rows:
        worksheet.append([excel_value(row.get(header)) for header in headers])


def unique_sheet_title(workbook: Workbook, raw_title: str) -> str:
    base = raw_title[:31]
    title = base
    suffix = 2
    while title in workbook.sheetnames:
        suffix_text = f"_{suffix}"
        title = f"{base[: 31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    return title


def write_report(
    config: RunConfig,
    *,
    status: str,
    database_name: str,
    database_host: str,
    variations: Sequence[dict[str, Any]],
    before_counts: dict[str, int],
    deleted_counts: dict[str, int] | None,
    after_counts: dict[str, int] | None,
    dependencies: Sequence[dict[str, Any]],
    related_impacts: Sequence[RelatedImpact],
    snapshots: Sequence[TableSnapshot],
    message: str,
    allow_overwrite: bool = False,
) -> None:
    output_path = config.output_path
    if output_path.exists() and not allow_overwrite:
        raise FileExistsError(f"File báo cáo đã tồn tại, không ghi đè: {output_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    summary.append(["Field", "Value"])
    summary_rows = [
        ("status", status),
        ("created_at", datetime.now().isoformat(timespec="seconds")),
        ("database", database_name),
        ("host", database_host),
        ("project_id", config.project_id),
        ("phase_id", config.phase_id),
        ("variation_ids", ", ".join(map(str, config.variation_ids))),
        ("execute_requested", config.execute),
        ("assume_yes", config.assume_yes),
        ("message", message),
    ]
    for field, value in summary_rows:
        summary.append([field, excel_value(value)])
    style_sheet(summary)

    count_sheet = workbook.create_sheet("Table counts")
    count_sheet.append(["table_name", "before", "deleted", "after"])
    for table_name in TARGET_TABLES:
        count_sheet.append(
            [
                table_name,
                before_counts.get(table_name, 0),
                None if deleted_counts is None else deleted_counts.get(table_name, 0),
                None if after_counts is None else after_counts.get(table_name, 0),
            ]
        )
    count_sheet.append(
        [
            "TOTAL",
            sum(before_counts.values()),
            None if deleted_counts is None else sum(deleted_counts.values()),
            None if after_counts is None else sum(after_counts.values()),
        ]
    )
    style_sheet(count_sheet)

    variation_sheet = workbook.create_sheet("Variations")
    append_mapping_rows(variation_sheet, list(variations))
    style_sheet(variation_sheet)

    dependency_sheet = workbook.create_sheet("FK dependencies")
    append_mapping_rows(dependency_sheet, list(dependencies))
    style_sheet(dependency_sheet)

    impact_sheet = workbook.create_sheet("Related impacts")
    append_mapping_rows(
        impact_sheet,
        [
            {
                "constraint_name": impact.constraint_name,
                "table_name": impact.table_name,
                "parent_table": impact.parent_table,
                "delete_rule": impact.delete_rule,
                "action": impact.action,
                "estimated_rows": impact.row_count,
                "note": impact.note,
            }
            for impact in related_impacts
        ],
    )
    style_sheet(impact_sheet)

    for snapshot in snapshots:
        worksheet = workbook.create_sheet(
            unique_sheet_title(workbook, snapshot.table_name)
        )
        worksheet.append(snapshot.columns)
        for row in snapshot.rows:
            worksheet.append([excel_value(value) for value in row])
        if not snapshot.rows:
            worksheet.append(["Không có dòng phù hợp"])
        if snapshot.truncated:
            worksheet.append([])
            worksheet.append(
                [
                    f"Chỉ xuất {len(snapshot.rows)}/{snapshot.total_rows} dòng; "
                    "xem Summary để biết tổng chính xác."
                ]
            )
        style_sheet(worksheet)

    temp_path = output_path.with_name(f".{output_path.name}.tmp.xlsx")
    workbook.save(temp_path)
    temp_path.replace(output_path)


def confirmation_text(config: RunConfig, database_name: str) -> str:
    variation_text = ",".join(map(str, config.variation_ids))
    return (
        f"DELETE {database_name} PROJECT {config.project_id} "
        f"PHASE {config.phase_id} VARIATIONS {variation_text}"
    )


def print_separator() -> None:
    print("=" * 78)


def main(argv: Sequence[str] | None = None) -> int:
    config = parse_args(argv)
    database_name, database_host = database_identity()

    variations: list[dict[str, Any]] = []
    before_counts: dict[str, int] = {}
    dependencies: list[dict[str, Any]] = []
    related_impacts: list[RelatedImpact] = []
    snapshots: list[TableSnapshot] = []
    deleted_counts: dict[str, int] | None = None
    after_counts: dict[str, int] | None = None
    report_written = False
    mutation_committed = False

    try:
        with connection.cursor() as cursor:
            validate_schema(cursor)
            validate_project_and_phase(cursor, config.project_id, config.phase_id)
            variations = fetch_variations(cursor, config.variation_ids)
            missing_master_ids = [
                item["variation_id"]
                for item in variations
                if not item["exists_in_master"]
            ]
            if missing_master_ids:
                raise ValueError(
                    "variation_id không tồn tại trong bảng variation: "
                    f"{missing_master_ids}"
                )

            used_ids = fetch_used_variation_ids(
                cursor,
                config.project_id,
                config.phase_id,
                config.variation_ids,
            )
            missing_target_ids = [
                variation_id
                for variation_id in config.variation_ids
                if variation_id not in used_ids
            ]
            if missing_target_ids:
                raise ValueError(
                    "Không tìm thấy variation_id trong project/phase ở các bảng "
                    f"được quản lý: {missing_target_ids}"
                )

            before_counts = count_target_rows(
                cursor,
                config.project_id,
                config.phase_id,
                config.variation_ids,
            )
            dependencies = fetch_fk_dependencies(cursor)
            related_impacts = fetch_related_impacts(
                cursor,
                dependencies,
                config.project_id,
                config.phase_id,
                config.variation_ids,
            )
            related_snapshot_tables = sorted(
                {
                    impact.table_name
                    for impact in related_impacts
                    if impact.row_count is not None and impact.row_count > 0
                }
            )
            snapshots = [
                fetch_table_snapshot(
                    cursor,
                    table_name,
                    config.project_id,
                    config.phase_id,
                    config.variation_ids,
                    config.report_row_limit,
                )
                for table_name in (*TARGET_TABLES, *related_snapshot_tables)
            ]

        print_separator()
        print("XEM TRƯỚC XÓA DỮ LIỆU VARIATION")
        print_separator()
        print(f"Database      : {database_name}")
        print(f"Host          : {database_host}")
        print(f"Project ID    : {config.project_id}")
        print(f"Phase ID      : {config.phase_id}")
        print(f"Variation IDs : {list(config.variation_ids)}")
        print(f"Execute       : {config.execute}")
        print()
        for table_name in TARGET_TABLES:
            print(f"{table_name:<32}: {before_counts[table_name]:>8} dòng")
        print(f"{'TỔNG':<32}: {sum(before_counts.values()):>8} dòng")
        nonzero_related_impacts = [
            impact
            for impact in related_impacts
            if impact.row_count is None or impact.row_count > 0
        ]
        if nonzero_related_impacts:
            print()
            print("Ảnh hưởng FK ngoài 4 bảng chính:")
            for impact in nonzero_related_impacts:
                count_text = (
                    "không xác định"
                    if impact.row_count is None
                    else str(impact.row_count)
                )
                print(f"- {impact.table_name}: {count_text} dòng; {impact.action}")

        write_report(
            config,
            status="PREVIEW",
            database_name=database_name,
            database_host=database_host,
            variations=variations,
            before_counts=before_counts,
            deleted_counts=None,
            after_counts=None,
            dependencies=dependencies,
            related_impacts=related_impacts,
            snapshots=snapshots,
            message="Đã xem trước; chưa xóa dữ liệu.",
        )
        report_written = True
        print(f"Báo cáo Excel : {config.output_path}")

        if not config.execute:
            print_separator()
            print("RUN_EXECUTE=False/không có --execute: chưa xóa dữ liệu.")
            return 0

        unauditable_impacts = [
            impact for impact in related_impacts if impact.row_count is None
        ]
        blocking_impacts = [
            impact
            for impact in related_impacts
            if impact.action == "POTENTIAL BLOCK" and (impact.row_count or 0) > 0
        ]
        if unauditable_impacts or blocking_impacts:
            reasons = []
            if unauditable_impacts:
                reasons.append(
                    "có quan hệ FK không thể ước lượng: "
                    + ", ".join(
                        sorted({impact.table_name for impact in unauditable_impacts})
                    )
                )
            if blocking_impacts:
                reasons.append(
                    "có dữ liệu RESTRICT/NO ACTION: "
                    + ", ".join(
                        sorted({impact.table_name for impact in blocking_impacts})
                    )
                )
            blocked_message = "Không cho phép execute vì " + "; ".join(reasons)
            write_report(
                config,
                status="BLOCKED",
                database_name=database_name,
                database_host=database_host,
                variations=variations,
                before_counts=before_counts,
                deleted_counts=None,
                after_counts=None,
                dependencies=dependencies,
                related_impacts=related_impacts,
                snapshots=snapshots,
                message=blocked_message,
                allow_overwrite=True,
            )
            print_separator()
            print(blocked_message)
            return 1

        expected_confirmation = confirmation_text(config, database_name)
        if not config.assume_yes:
            print_separator()
            print("Để xác nhận xóa, nhập chính xác nội dung sau:")
            print(expected_confirmation)
            actual_confirmation = input("> ").strip()
            if actual_confirmation != expected_confirmation:
                write_report(
                    config,
                    status="CANCELLED",
                    database_name=database_name,
                    database_host=database_host,
                    variations=variations,
                    before_counts=before_counts,
                    deleted_counts=None,
                    after_counts=None,
                    dependencies=dependencies,
                    related_impacts=related_impacts,
                    snapshots=snapshots,
                    message="Người dùng hủy hoặc nhập sai chuỗi xác nhận.",
                    allow_overwrite=True,
                )
                print("Đã hủy. Không có dữ liệu nào bị xóa.")
                return 0

        delete_order = order_tables_child_first(TARGET_TABLES, dependencies)
        deleted_counts, after_counts = delete_target_rows(
            config, before_counts, delete_order
        )
        mutation_committed = True
        write_report(
            config,
            status="DELETED",
            database_name=database_name,
            database_host=database_host,
            variations=variations,
            before_counts=before_counts,
            deleted_counts=deleted_counts,
            after_counts=after_counts,
            dependencies=dependencies,
            related_impacts=related_impacts,
            snapshots=snapshots,
            message="Xóa thành công trong một transaction và đã kiểm tra không còn dòng.",
            allow_overwrite=True,
        )

        print_separator()
        print("XÓA THÀNH CÔNG")
        print_separator()
        for table_name in delete_order:
            print(f"{table_name:<32}: {deleted_counts[table_name]:>8} dòng")
        print(f"Báo cáo Excel : {config.output_path}")
        return 0

    except KeyboardInterrupt:
        if mutation_committed:
            message = (
                "Database đã commit xóa; chương trình bị ngắt khi cập nhật báo cáo."
            )
            interrupt_status = "INTERRUPTED_AFTER_DELETE"
        else:
            message = "Người dùng ngắt chương trình; transaction (nếu có) đã rollback."
            interrupt_status = "INTERRUPTED"
        if report_written:
            try:
                write_report(
                    config,
                    status=interrupt_status,
                    database_name=database_name,
                    database_host=database_host,
                    variations=variations,
                    before_counts=before_counts,
                    deleted_counts=deleted_counts,
                    after_counts=after_counts,
                    dependencies=dependencies,
                    related_impacts=related_impacts,
                    snapshots=snapshots,
                    message=message,
                    allow_overwrite=True,
                )
            except Exception as report_error:
                print(f"Không thể cập nhật báo cáo: {report_error}")
        print(f"\n{message}")
        return 130
    except Exception as error:
        message = f"{type(error).__name__}: {error}"
        report_status = "FAILED_AFTER_DELETE" if mutation_committed else "FAILED"
        if mutation_committed:
            message = (
                "Database đã commit xóa nhưng cập nhật báo cáo sau xóa thất bại. "
                + message
            )
        if report_written or not config.output_path.exists():
            try:
                write_report(
                    config,
                    status=report_status,
                    database_name=database_name,
                    database_host=database_host,
                    variations=variations,
                    before_counts=before_counts,
                    deleted_counts=deleted_counts,
                    after_counts=after_counts,
                    dependencies=dependencies,
                    related_impacts=related_impacts,
                    snapshots=snapshots,
                    message=message,
                    allow_overwrite=report_written,
                )
            except Exception as report_error:
                print(f"Không thể cập nhật báo cáo lỗi: {report_error}")
        print_separator()
        print("XẢY RA LỖI")
        print_separator()
        print(message)
        return 1
    finally:
        connection.close()


if __name__ == "__main__":
    raise SystemExit(main())
