"""Preview or safely import SE master rows from Excel.

The configured worksheet maps ``WP`` to ``wp.wp``, ``Category`` to
``se_parameter.parameter_name_1``, ``パラメータ名`` to
``se_parameter.parameter_name_2``, and ``単位`` to ``parameter_unit``.
When duplicate ``wp.wp`` or ``units.unit`` values exist, the matching record
with the smallest ``id`` is selected and recorded in the report. A blank
``単位`` value is resolved against a row where ``units.unit IS NULL`` instead
of inserting ``se_parameter.unit_id`` as ``NULL``.
New ``se_parameter.id`` values are allocated from ``MAX(id) + 1`` and every
insert uses ``hierarchy_id=1``. Missing WPs and units (including the NULL-unit
reference used for blank cells) are added in the same transaction so their
foreign keys remain usable by existing joins.

Every run writes an Excel report. The default is preview-only; database
mutation requires an explicit execute flag and exact confirmation.
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402

# PyCharm Run-button configuration. Always preview before enabling execution.
RUN_EXCEL_PATH = Path("/Users/TrongKieu/Downloads/seリストへの追加.xlsx")
RUN_SHEET_NAME = "新規追加する項目"
RUN_HIERARCHY_ID = 1
RUN_CREATE_MISSING_WPS = False
RUN_CREATE_MISSING_UNITS = False
RUN_REPORT_PATH: Path | None = None
RUN_EXECUTE = False
RUN_ASSUME_YES = False

DEFAULT_REPORT_DIRECTORY = PROJECT_ROOT / "outputs" / "se_master_add_reports"
EXPECTED_HEADERS = (
    "Se-paramemter-ID",
    "WP",
    "Category",
    "パラメータ名",
    "単位",
)
PREVIEW_ROW_LIMIT = 100


@dataclass
class PlanRow:
    excel_row: int
    source_action: str | None
    wp: str | None
    category: str | None
    parameter_name: str | None
    parameter_unit: str | None
    hierarchy_id: int
    wp_id: int | None = None
    unit_id: int | None = None
    se_parameter_id: int | None = None
    action: str = "PENDING"
    result: str = "PENDING"
    message: str = ""
    reference_notes: list[str] = field(default_factory=list)


@dataclass
class UnitPlan:
    unit_id: int
    unit: str | None
    source_rows: list[int] = field(default_factory=list)
    result: str = "WOULD_INSERT"


@dataclass
class WpPlan:
    wp_id: int
    wp: str
    source_rows: list[int] = field(default_factory=list)
    result: str = "WOULD_INSERT"


@dataclass
class ImportPlan:
    rows: list[PlanRow]
    units: list[UnitPlan]
    hierarchy_name: str | None = None
    wps: list[WpPlan] = field(default_factory=list)

    @property
    def errors(self) -> list[PlanRow]:
        return [row for row in self.rows if row.result == "ERROR"]

    @property
    def insert_rows(self) -> list[PlanRow]:
        return [row for row in self.rows if row.result == "WOULD_INSERT"]

    @property
    def mutation_count(self) -> int:
        return len(self.insert_rows) + len(
            [unit for unit in self.units if unit.result == "WOULD_INSERT"]
        ) + len([wp for wp in self.wps if wp.result == "WOULD_INSERT"])


@dataclass(frozen=True)
class RunResult:
    status: str
    report_path: Path
    inserted_count: int
    plan: ImportPlan


class ToolRunError(RuntimeError):
    """Raised after a failure report has been written."""

    def __init__(self, message: str, report_path: Path) -> None:
        super().__init__(message)
        self.report_path = report_path


def normalize_header(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value)).strip().casefold()
    return "".join(character for character in text if not character.isspace())


def normalize_text(value: object) -> str | None:
    if value is None:
        return None
    cleaned = str(value).strip()
    return cleaned or None


def normalize_match_key(value: object) -> str | None:
    cleaned = normalize_text(value)
    if cleaned is None:
        return None
    return unicodedata.normalize("NFKC", cleaned)


def normalize_positive_id(value: object, *, label: str) -> int:
    if isinstance(value, bool):
        raise ValueError(f"{label} must be a positive integer, not boolean.")
    if isinstance(value, float) and not value.is_integer():
        raise ValueError(f"{label} must be an integer: {value!r}")
    try:
        normalized = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid {label}: {value!r}") from exc
    if normalized <= 0:
        raise ValueError(f"{label} must be greater than zero: {normalized}")
    return normalized


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create SQL placeholders for an empty sequence.")
    return ", ".join(["%s"] * len(values))


def safe_excel_value(value: object) -> object:
    """Prevent report text from being interpreted as an Excel formula."""
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone().replace(tzinfo=None)
        return value
    if value is None or isinstance(value, (bool, int, float)):
        return value
    text = str(value)
    if text.startswith(("=", "+", "-", "@")):
        return f"'{text}"
    return text


def make_report_path(report_path: Path | str | None = None) -> Path:
    if report_path is not None:
        path = Path(report_path).expanduser().resolve()
        if path.suffix.lower() != ".xlsx":
            raise ValueError("Report path must use the .xlsx extension.")
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    DEFAULT_REPORT_DIRECTORY.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S_%f")
    return DEFAULT_REPORT_DIRECTORY / f"se_master_add_{timestamp}.xlsx"


def read_excel_rows(
    excel_path: Path | str,
    sheet_name: str,
    *,
    hierarchy_id: int,
) -> list[PlanRow]:
    """Read and validate the five-column source worksheet."""
    path = Path(excel_path).expanduser().resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Excel file was not found: {path}")
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Excel input must be an .xlsx or .xlsm file.")

    normalized_hierarchy_id = normalize_positive_id(hierarchy_id, label="hierarchy_id")
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Sheet {sheet_name!r} was not found; available={workbook.sheetnames}"
            )
        worksheet = workbook[sheet_name]
        actual_headers = tuple(
            worksheet.cell(1, column).value
            for column in range(1, len(EXPECTED_HEADERS) + 1)
        )
        for column, (actual, expected) in enumerate(
            zip(actual_headers, EXPECTED_HEADERS, strict=True), start=1
        ):
            if normalize_header(actual) != normalize_header(expected):
                letter = get_column_letter(column)
                raise ValueError(
                    f"Expected header {expected!r} in {letter}1, found {actual!r}."
                )

        rows: list[PlanRow] = []
        for row_number in range(2, worksheet.max_row + 1):
            cells = [worksheet.cell(row_number, column) for column in range(1, 6)]
            raw_values = [cell.value for cell in cells]
            if all(value in (None, "") for value in raw_values):
                continue

            row = PlanRow(
                excel_row=row_number,
                source_action=normalize_text(raw_values[0]),
                wp=normalize_text(raw_values[1]),
                category=normalize_text(raw_values[2]),
                parameter_name=normalize_text(raw_values[3]),
                parameter_unit=normalize_text(raw_values[4]),
                hierarchy_id=normalized_hierarchy_id,
            )
            rows.append(row)
            formula_columns = [
                get_column_letter(index)
                for index, cell in enumerate(cells[1:], start=2)
                if cell.data_type == "f"
            ]
            errors: list[str] = []
            if formula_columns:
                errors.append(
                    "Formulas are not allowed in columns "
                    + ", ".join(formula_columns)
                    + "."
                )
            if row.wp is None:
                errors.append("WP is blank.")
            if row.category is None:
                errors.append("Category is blank.")
            if errors:
                row.action = "ERROR"
                row.result = "ERROR"
                row.message = " ".join(errors)

        if not rows:
            raise ValueError("No input rows were found in the worksheet.")
        return rows
    finally:
        workbook.close()


def database_identity(db_connection=connection) -> tuple[str, str, str]:
    settings = db_connection.settings_dict
    return (
        str(settings.get("NAME") or "<unknown>"),
        str(settings.get("HOST") or "localhost"),
        str(settings.get("PORT") or "3306"),
    )


def _catalog_by_name(
    rows: Sequence[tuple[object, object]],
) -> dict[str, list[tuple[int, str]]]:
    catalog: dict[str, list[tuple[int, str]]] = {}
    for raw_id, raw_name in rows:
        normalized_name = normalize_match_key(raw_name)
        if normalized_name is None:
            continue
        catalog.setdefault(normalized_name, []).append((int(raw_id), str(raw_name)))
    for matches in catalog.values():
        matches.sort(key=lambda match: match[0])
    return catalog


def _select_smallest_id(
    matches: Sequence[tuple[int, object]],
) -> tuple[int, list[int]]:
    """Return the smallest ID and all candidate IDs in deterministic order."""
    if not matches:
        raise ValueError("Cannot select an ID from an empty match list.")
    matching_ids = sorted({match_id for match_id, _ in matches})
    return matching_ids[0], matching_ids


def _message_with_reference_notes(row: PlanRow, message: str) -> str:
    if not row.reference_notes:
        return message
    return f"{message} {' '.join(row.reference_notes)}"


def business_key(
    *,
    hierarchy_id: int,
    wp_id: int,
    category: str | None,
    parameter_name: str | None,
    parameter_unit: str | None,
    unit_id: int | None,
) -> tuple[object, ...]:
    return (
        hierarchy_id,
        wp_id,
        normalize_match_key(category),
        normalize_match_key(parameter_name),
        normalize_match_key(parameter_unit),
        unit_id,
    )


def _next_id(cursor: Any, table_name: str, *, lock: bool) -> int:
    if table_name not in {"se_parameter", "units", "wp"}:
        raise ValueError(f"Unsupported ID table: {table_name!r}")
    lock_clause = " FOR UPDATE" if lock else ""
    cursor.execute(
        f"SELECT id FROM `{table_name}` ORDER BY id DESC LIMIT 1{lock_clause}"
    )
    row = cursor.fetchone()
    return (int(row[0]) if row else 0) + 1


def build_plan(
    cursor: Any,
    rows: list[PlanRow],
    *,
    hierarchy_id: int,
    create_missing_wps: bool,
    create_missing_units: bool,
    lock_ids: bool = False,
) -> ImportPlan:
    """Resolve references, existing records, and collision-free planned IDs."""
    hierarchy_id = normalize_positive_id(hierarchy_id, label="hierarchy_id")
    cursor.execute("SELECT hierarchy FROM hierarchy WHERE id = %s", [hierarchy_id])
    hierarchy_record = cursor.fetchone()
    if hierarchy_record is None:
        message = f"hierarchy.id={hierarchy_id} does not exist."
        for row in rows:
            if row.result == "PENDING":
                row.action = "ERROR"
                row.result = "ERROR"
                row.message = message
        return ImportPlan(rows, [], None)
    hierarchy_name = hierarchy_record[0]

    cursor.execute("SELECT id, wp FROM wp ORDER BY id")
    wp_catalog = _catalog_by_name(cursor.fetchall())
    cursor.execute("SELECT id, unit FROM units ORDER BY id")
    unit_records = cursor.fetchall()
    unit_catalog = _catalog_by_name(unit_records)
    null_unit_matches = [
        (int(unit_id), unit)
        for unit_id, unit in unit_records
        if unit is None
    ]

    next_wp_id = _next_id(cursor, "wp", lock=lock_ids)
    next_unit_id = _next_id(cursor, "units", lock=lock_ids)
    next_parameter_id = _next_id(cursor, "se_parameter", lock=lock_ids)
    new_wps_by_name: dict[str, WpPlan] = {}
    new_units_by_name: dict[str | None, UnitPlan] = {}

    for row in rows:
        if row.result != "PENDING":
            continue
        wp_key = normalize_match_key(row.wp)
        assert wp_key is not None
        wp_matches = wp_catalog.get(wp_key, [])
        if wp_matches:
            row.wp_id, matching_wp_ids = _select_smallest_id(wp_matches)
            if len(matching_wp_ids) > 1:
                row.reference_notes.append(
                    "WP matched multiple IDs "
                    f"({', '.join(str(value) for value in matching_wp_ids)}); "
                    f"selected smallest ID {row.wp_id}."
                )
        elif not create_missing_wps:
            row.action = "ERROR"
            row.result = "ERROR"
            row.message = "WP was not found in wp.wp."
            continue
        else:
            wp_plan = new_wps_by_name.get(wp_key)
            if wp_plan is None:
                wp_plan = WpPlan(next_wp_id, row.wp or "")
                new_wps_by_name[wp_key] = wp_plan
                next_wp_id += 1
            wp_plan.source_rows.append(row.excel_row)
            row.wp_id = wp_plan.wp_id
            row.reference_notes.append(
                f"WP was not found; planned new wp.id {row.wp_id}."
            )

        unit_key = normalize_match_key(row.parameter_unit)
        unit_matches = (
            null_unit_matches
            if unit_key is None
            else unit_catalog.get(unit_key, [])
        )
        if unit_matches:
            row.unit_id, matching_unit_ids = _select_smallest_id(unit_matches)
            if len(matching_unit_ids) > 1:
                unit_label = "NULL unit" if unit_key is None else "Unit"
                row.reference_notes.append(
                    f"{unit_label} matched multiple IDs "
                    f"({', '.join(str(value) for value in matching_unit_ids)}); "
                    f"selected smallest ID {row.unit_id}."
                )
        elif not create_missing_units:
            row.action = "ERROR"
            row.result = "ERROR"
            row.message = _message_with_reference_notes(
                row,
                (
                    "NULL unit was not found in units.unit."
                    if unit_key is None
                    else "Unit was not found in units.unit."
                ),
            )
        else:
            unit_plan = new_units_by_name.get(unit_key)
            if unit_plan is None:
                unit_plan = UnitPlan(next_unit_id, row.parameter_unit)
                new_units_by_name[unit_key] = unit_plan
                next_unit_id += 1
            unit_plan.source_rows.append(row.excel_row)
            row.unit_id = unit_plan.unit_id
            if unit_key is None:
                row.reference_notes.append(
                    f"NULL unit was not found; planned new units.id {row.unit_id}."
                )

    valid_wp_ids = sorted(
        {row.wp_id for row in rows if row.result == "PENDING" and row.wp_id is not None}
    )
    existing_by_key: dict[tuple[object, ...], list[int]] = {}
    if valid_wp_ids:
        placeholders = make_placeholders(valid_wp_ids)
        lock_clause = " FOR UPDATE" if lock_ids else ""
        cursor.execute(
            f"""
            SELECT id, hierarchy_id, parameter_name_1, parameter_name_2,
                   parameter_unit, wp_id, unit_id
            FROM se_parameter
            WHERE hierarchy_id = %s
              AND wp_id IN ({placeholders})
            ORDER BY id{lock_clause}
            """,
            [hierarchy_id, *valid_wp_ids],
        )
        for (
            se_parameter_id,
            existing_hierarchy_id,
            category,
            parameter_name,
            parameter_unit,
            wp_id,
            unit_id,
        ) in cursor.fetchall():
            key = business_key(
                hierarchy_id=int(existing_hierarchy_id),
                wp_id=int(wp_id),
                category=category,
                parameter_name=parameter_name,
                parameter_unit=parameter_unit,
                unit_id=unit_id,
            )
            existing_by_key.setdefault(key, []).append(int(se_parameter_id))

    staged_by_key: dict[tuple[object, ...], int] = {}
    first_source_row_by_key: dict[tuple[object, ...], int] = {}
    for row in rows:
        if row.result != "PENDING":
            continue
        assert row.wp_id is not None
        key = business_key(
            hierarchy_id=hierarchy_id,
            wp_id=row.wp_id,
            category=row.category,
            parameter_name=row.parameter_name,
            parameter_unit=row.parameter_unit,
            unit_id=row.unit_id,
        )
        existing_ids = existing_by_key.get(key, [])
        if existing_ids:
            row.se_parameter_id = min(existing_ids)
            row.action = "SKIP_EXISTING"
            row.result = "EXISTING"
            row.message = (
                "Business key already exists in se_parameter. Matching IDs: "
                + ", ".join(str(value) for value in existing_ids)
            )
            row.message = _message_with_reference_notes(row, row.message)
            continue
        if key in staged_by_key:
            row.se_parameter_id = staged_by_key[key]
            row.action = "SKIP_DUPLICATE_INPUT"
            row.result = "DUPLICATE_INPUT"
            row.message = (
                "Same business key already appeared at Excel row "
                f"{first_source_row_by_key[key]}."
            )
            row.message = _message_with_reference_notes(row, row.message)
            continue

        row.se_parameter_id = next_parameter_id
        next_parameter_id += 1
        staged_by_key[key] = row.se_parameter_id
        first_source_row_by_key[key] = row.excel_row
        row.action = "INSERT"
        row.result = "WOULD_INSERT"
        row.message = _message_with_reference_notes(
            row, "Preview only; no database row has been inserted."
        )

    used_new_unit_ids = {
        row.unit_id
        for row in rows
        if row.result in {"WOULD_INSERT", "DUPLICATE_INPUT"} and row.unit_id is not None
    }
    used_new_wp_ids = {
        row.wp_id
        for row in rows
        if row.result in {"WOULD_INSERT", "DUPLICATE_INPUT"} and row.wp_id is not None
    }
    wp_plans = [
        wp for wp in new_wps_by_name.values() if wp.wp_id in used_new_wp_ids
    ]
    unit_plans = [
        unit for unit in new_units_by_name.values() if unit.unit_id in used_new_unit_ids
    ]

    if lock_ids:
        new_wp_ids = {wp.wp_id for wp in wp_plans}
        existing_wp_ids = sorted(
            {
                row.wp_id
                for row in rows
                if row.wp_id is not None and row.wp_id not in new_wp_ids
            }
        )
        if existing_wp_ids:
            placeholders = make_placeholders(existing_wp_ids)
            cursor.execute(
                f"SELECT id FROM wp WHERE id IN ({placeholders}) FOR UPDATE",
                existing_wp_ids,
            )
            locked_wp_ids = {int(record[0]) for record in cursor.fetchall()}
            if locked_wp_ids != set(existing_wp_ids):
                missing = sorted(set(existing_wp_ids) - locked_wp_ids)
                raise RuntimeError(f"WP rows disappeared before import: {missing}")

    return ImportPlan(rows, unit_plans, hierarchy_name, wp_plans)


def plan_signature(plan: ImportPlan) -> tuple[tuple[object, ...], ...]:
    wp_items = tuple(
        ("wp", wp.wp_id, wp.wp)
        for wp in plan.wps
        if wp.result == "WOULD_INSERT"
    )
    unit_items = tuple(
        ("units", unit.unit_id, unit.unit)
        for unit in plan.units
        if unit.result == "WOULD_INSERT"
    )
    parameter_items = tuple(
        (
            "se_parameter",
            row.se_parameter_id,
            row.hierarchy_id,
            row.wp_id,
            row.category,
            row.parameter_name,
            row.parameter_unit,
            row.unit_id,
        )
        for row in plan.rows
        if row.result == "WOULD_INSERT"
    )
    return wp_items + unit_items + parameter_items


def execute_plan(cursor: Any, plan: ImportPlan) -> int:
    """Insert a validated, locked plan and verify all new IDs."""
    if plan.errors:
        raise RuntimeError("Execution plan contains validation errors.")

    inserted_wp_ids: list[int] = []
    for wp in plan.wps:
        if wp.result != "WOULD_INSERT":
            continue
        cursor.execute(
            "INSERT INTO wp (id, wp) VALUES (%s, %s)",
            [wp.wp_id, wp.wp],
        )
        if cursor.rowcount != 1:
            raise RuntimeError(
                f"Unexpected insert count for wp.id={wp.wp_id}: {cursor.rowcount}"
            )
        inserted_wp_ids.append(wp.wp_id)
        wp.result = "INSERTED"

    inserted_unit_ids: list[int] = []
    for unit in plan.units:
        if unit.result != "WOULD_INSERT":
            continue
        cursor.execute(
            "INSERT INTO units (id, unit) VALUES (%s, %s)",
            [unit.unit_id, unit.unit],
        )
        if cursor.rowcount != 1:
            raise RuntimeError(
                f"Unexpected insert count for units.id={unit.unit_id}: "
                f"{cursor.rowcount}"
            )
        inserted_unit_ids.append(unit.unit_id)
        unit.result = "INSERTED"

    inserted_parameter_ids: list[int] = []
    for row in plan.rows:
        if row.result != "WOULD_INSERT":
            continue
        assert row.se_parameter_id is not None
        assert row.wp_id is not None
        cursor.execute(
            """
            INSERT INTO se_parameter (
                id, hierarchy_id, parameter_name_1, parameter_name_2,
                parameter_unit, wp_id, unit_id
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            """,
            [
                row.se_parameter_id,
                row.hierarchy_id,
                row.category,
                row.parameter_name,
                row.parameter_unit,
                row.wp_id,
                row.unit_id,
            ],
        )
        if cursor.rowcount != 1:
            raise RuntimeError(
                f"Unexpected insert count for se_parameter.id="
                f"{row.se_parameter_id}: {cursor.rowcount}"
            )
        inserted_parameter_ids.append(row.se_parameter_id)
        row.result = "INSERTED"
        row.message = _message_with_reference_notes(
            row, "Inserted in the committed transaction."
        )

    for table_name, inserted_ids in (
        ("wp", inserted_wp_ids),
        ("units", inserted_unit_ids),
        ("se_parameter", inserted_parameter_ids),
    ):
        if not inserted_ids:
            continue
        placeholders = make_placeholders(inserted_ids)
        cursor.execute(
            f"SELECT COUNT(*) FROM `{table_name}` " f"WHERE id IN ({placeholders})",
            inserted_ids,
        )
        actual_count = int(cursor.fetchone()[0])
        if actual_count != len(inserted_ids):
            raise RuntimeError(
                f"Post-insert verification failed for {table_name}: "
                f"found {actual_count}, expected {len(inserted_ids)}."
            )
    return len(inserted_parameter_ids)


def confirmation_text(database: str, plan: ImportPlan) -> str:
    se_ids = ",".join(
        str(row.se_parameter_id) for row in plan.rows if row.result == "WOULD_INSERT"
    )
    unit_ids = ",".join(
        str(unit.unit_id) for unit in plan.units if unit.result == "WOULD_INSERT"
    )
    wp_ids = ",".join(
        str(wp.wp_id) for wp in plan.wps if wp.result == "WOULD_INSERT"
    )
    return (
        f"IMPORT {database} SE_PARAMETER IDS {se_ids or '-'} "
        f"WP IDS {wp_ids or '-'} "
        f"UNITS IDS {unit_ids or '-'}"
    )


def confirm_execution(
    database: str,
    plan: ImportPlan,
    *,
    assume_yes: bool,
) -> bool:
    expected = confirmation_text(database, plan)
    if assume_yes:
        print("Confirmation bypass explicitly enabled with --yes/RUN_ASSUME_YES.")
        return True
    print("Database mutation requires exact confirmation.")
    print(f"Type exactly: {expected}")
    return input("> ").strip() == expected


def print_plan(
    plan: ImportPlan,
    *,
    database: str,
    host: str,
    port: str,
    excel_path: Path,
    sheet_name: str,
    report_path: Path,
    execute: bool,
) -> None:
    print("SE master import plan")
    print(f"Database: {database}")
    print(f"Host: {host}:{port}")
    print(f"Input Excel: {excel_path}")
    print(f"Input sheet: {sheet_name}")
    print(
        f"Hierarchy: {plan.rows[0].hierarchy_id if plan.rows else '-'} "
        f"({plan.hierarchy_name!r})"
    )
    print(f"Mode: {'EXECUTE' if execute else 'PREVIEW'}")
    print(f"Report: {report_path}")
    print("Affected tables and estimated new rows:")
    print(f"- wp: {len(plan.wps)}")
    print(f"- units: {len(plan.units)}")
    print(f"- se_parameter: {len(plan.insert_rows)}")
    print(f"Existing rows skipped: {sum(r.result == 'EXISTING' for r in plan.rows)}")
    print(
        "Duplicate input rows skipped: "
        f"{sum(r.result == 'DUPLICATE_INPUT' for r in plan.rows)}"
    )
    print(f"Validation errors: {len(plan.errors)}")
    print("Selected records:")
    for row in plan.rows[:PREVIEW_ROW_LIMIT]:
        print(
            f"- [{row.result}] Excel row {row.excel_row}: "
            f"se_parameter.id={row.se_parameter_id!r}; wp={row.wp!r} "
            f"(id={row.wp_id!r}); category={row.category!r}; "
            f"name={row.parameter_name!r}; unit={row.parameter_unit!r} "
            f"(id={row.unit_id!r})"
        )
    if len(plan.rows) > PREVIEW_ROW_LIMIT:
        print(f"- ... {len(plan.rows) - PREVIEW_ROW_LIMIT} more rows")


def _style_sheet(sheet: Any, *, header_row: int = 1) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, color="1F1F1F")
    thin_gray = Side(style="thin", color="D9E2F3")
    for cell in sheet[header_row]:
        if cell.value is not None:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in sheet.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_gray)
    columns = list(sheet.columns)
    for column_index in range(1, sheet.max_column + 1):
        values = [
            "" if cell.value is None else str(cell.value)
            for cell in columns[column_index - 1]
        ]
        width = min(max(max((len(value) for value in values), default=0) + 2, 10), 45)
        sheet.column_dimensions[get_column_letter(column_index)].width = width
    sheet.sheet_view.showGridLines = False


def write_report(
    report_path: Path,
    plan: ImportPlan,
    *,
    status: str,
    database: str,
    host: str,
    port: str,
    excel_path: Path,
    sheet_name: str,
    execute_requested: bool,
    assume_yes: bool,
    create_missing_wps: bool,
    create_missing_units: bool,
    started_at: datetime,
    completed_at: datetime,
    error: str | None = None,
) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Summary"
    results = workbook.create_sheet("Results")
    wps = workbook.create_sheet("WPs")
    units = workbook.create_sheet("Units")

    summary.append(["SE master import report"])
    summary["A1"].font = Font(name="Arial", size=16, bold=True, color="1F1F1F")
    summary.append([])
    summary.append(["Item", "Value"])
    summary_rows: list[tuple[str, object]] = [
        ("Status", status),
        ("Mode", "EXECUTE" if execute_requested else "PREVIEW"),
        ("Database", database),
        ("Host", f"{host}:{port}"),
        ("Target tables", "wp, units, se_parameter"),
        ("Input file", str(excel_path)),
        ("Input sheet", sheet_name),
        ("Hierarchy ID", plan.rows[0].hierarchy_id if plan.rows else None),
        ("Hierarchy", plan.hierarchy_name),
        ("Create missing WPs", create_missing_wps),
        ("Create missing units", create_missing_units),
        ("Non-interactive confirmation", assume_yes),
        ("Input rows", len(plan.rows)),
        ("SE rows planned", sum(r.action == "INSERT" for r in plan.rows)),
        ("SE rows inserted", sum(r.result == "INSERTED" for r in plan.rows)),
        ("Existing rows skipped", sum(r.result == "EXISTING" for r in plan.rows)),
        (
            "Duplicate input rows skipped",
            sum(r.result == "DUPLICATE_INPUT" for r in plan.rows),
        ),
        ("WPs planned", len(plan.wps)),
        ("WPs inserted", sum(wp.result == "INSERTED" for wp in plan.wps)),
        ("Units planned", len(plan.units)),
        ("Units inserted", sum(u.result == "INSERTED" for u in plan.units)),
        ("Error rows", len(plan.errors)),
        ("Started", started_at),
        ("Completed", completed_at),
        ("Error", error or ""),
    ]
    summary_row_numbers: dict[str, int] = {}
    for item, value in summary_rows:
        summary.append([item, safe_excel_value(value)])
        summary_row_numbers[item] = summary.max_row
    _style_sheet(summary, header_row=3)
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 90
    summary.cell(summary_row_numbers["Started"], 2).number_format = (
        "yyyy-mm-dd hh:mm:ss"
    )
    summary.cell(summary_row_numbers["Completed"], 2).number_format = (
        "yyyy-mm-dd hh:mm:ss"
    )
    status_colors = {
        "COMMITTED": "C6EFCE",
        "PREVIEW": "FFF2CC",
        "NO_CHANGES": "D9EAD3",
        "READY_TO_EXECUTE": "FCE5CD",
        "CANCELLED": "E7E6E6",
        "BLOCKED": "F4CCCC",
        "ROLLED_BACK": "F4CCCC",
    }
    summary["B4"].fill = PatternFill(
        "solid", fgColor=status_colors.get(status, "D9E2F3")
    )
    summary["B4"].font = Font(name="Arial", size=10, bold=True)

    result_headers = [
        "excel_row",
        "source_action",
        "se_parameter_id",
        "hierarchy_id",
        "wp",
        "wp_id",
        "category",
        "parameter_name",
        "parameter_unit",
        "unit_id",
        "action",
        "result",
        "message",
    ]
    results.append(result_headers)
    for row in plan.rows:
        results.append(
            [
                safe_excel_value(row.excel_row),
                safe_excel_value(row.source_action),
                safe_excel_value(row.se_parameter_id),
                safe_excel_value(row.hierarchy_id),
                safe_excel_value(row.wp),
                safe_excel_value(row.wp_id),
                safe_excel_value(row.category),
                safe_excel_value(row.parameter_name),
                safe_excel_value(row.parameter_unit),
                safe_excel_value(row.unit_id),
                safe_excel_value(row.action),
                safe_excel_value(row.result),
                safe_excel_value(row.message),
            ]
        )
    _style_sheet(results)
    results.freeze_panes = "A2"
    results.auto_filter.ref = results.dimensions
    fixed_widths = {
        "A": 12,
        "B": 18,
        "C": 18,
        "D": 14,
        "E": 22,
        "F": 12,
        "G": 35,
        "H": 45,
        "I": 18,
        "J": 12,
        "K": 22,
        "L": 22,
        "M": 70,
    }
    for column, width in fixed_widths.items():
        results.column_dimensions[column].width = width
    for row_number in range(2, results.max_row + 1):
        result_cell = results.cell(row_number, 12)
        if result_cell.value in {"ERROR", "BLOCKED", "ROLLED_BACK"}:
            result_cell.fill = PatternFill("solid", fgColor="F4CCCC")
        elif result_cell.value == "INSERTED":
            result_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        elif result_cell.value == "WOULD_INSERT":
            result_cell.fill = PatternFill("solid", fgColor="FFF2CC")

    wps.append(["wp_id", "wp", "source_rows", "result"])
    for wp in plan.wps:
        wps.append(
            [
                wp.wp_id,
                safe_excel_value(wp.wp),
                ", ".join(str(value) for value in wp.source_rows),
                wp.result,
            ]
        )
    _style_sheet(wps)
    wps.freeze_panes = "A2"
    wps.auto_filter.ref = wps.dimensions
    wps.column_dimensions["A"].width = 14
    wps.column_dimensions["B"].width = 24
    wps.column_dimensions["C"].width = 50
    wps.column_dimensions["D"].width = 22

    units.append(["unit_id", "unit", "source_rows", "result"])
    for unit in plan.units:
        units.append(
            [
                unit.unit_id,
                safe_excel_value(unit.unit),
                ", ".join(str(value) for value in unit.source_rows),
                unit.result,
            ]
        )
    _style_sheet(units)
    units.freeze_panes = "A2"
    units.auto_filter.ref = units.dimensions
    units.column_dimensions["A"].width = 14
    units.column_dimensions["B"].width = 24
    units.column_dimensions["C"].width = 50
    units.column_dimensions["D"].width = 22

    report_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{report_path.stem}_",
            suffix=".xlsx",
            dir=report_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)
        workbook.save(temporary_path)
        temporary_path.replace(report_path)
    finally:
        workbook.close()
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def mark_plan_blocked(plan: ImportPlan) -> None:
    for row in plan.rows:
        if row.result == "WOULD_INSERT":
            row.result = "BLOCKED"
            row.message = _message_with_reference_notes(
                row, "Not executed because another input row has an error."
            )
    for unit in plan.units:
        if unit.result == "WOULD_INSERT":
            unit.result = "BLOCKED"
    for wp in plan.wps:
        if wp.result == "WOULD_INSERT":
            wp.result = "BLOCKED"


def import_se_master(
    excel_path: Path | str,
    *,
    sheet_name: str = RUN_SHEET_NAME,
    hierarchy_id: int = RUN_HIERARCHY_ID,
    create_missing_wps: bool = RUN_CREATE_MISSING_WPS,
    create_missing_units: bool = RUN_CREATE_MISSING_UNITS,
    execute: bool = False,
    assume_yes: bool = False,
    report_path: Path | str | None = None,
    db_connection=connection,
) -> RunResult:
    """Preview or execute an idempotent SE master import."""
    if assume_yes and not execute:
        raise ValueError("--yes/RUN_ASSUME_YES requires execute mode.")

    input_path = Path(excel_path).expanduser().resolve()
    output_path = make_report_path(report_path)
    started_at = datetime.now().astimezone()
    rows: list[PlanRow] = []
    plan = ImportPlan(rows, [])
    database, host, port = database_identity(db_connection)

    try:
        rows = read_excel_rows(input_path, sheet_name, hierarchy_id=hierarchy_id)
        plan = ImportPlan(rows, [])
        with db_connection.cursor() as cursor:
            plan = build_plan(
                cursor,
                rows,
                hierarchy_id=hierarchy_id,
                create_missing_wps=create_missing_wps,
                create_missing_units=create_missing_units,
            )
    except Exception as exc:
        completed_at = datetime.now().astimezone()
        write_report(
            output_path,
            plan,
            status="BLOCKED",
            database=database,
            host=host,
            port=port,
            excel_path=input_path,
            sheet_name=sheet_name,
            execute_requested=execute,
            assume_yes=assume_yes,
            create_missing_wps=create_missing_wps,
            create_missing_units=create_missing_units,
            started_at=started_at,
            completed_at=completed_at,
            error=f"{type(exc).__name__}: {exc}",
        )
        raise ToolRunError(str(exc), output_path) from exc

    print_plan(
        plan,
        database=database,
        host=host,
        port=port,
        excel_path=input_path,
        sheet_name=sheet_name,
        report_path=output_path,
        execute=execute,
    )

    if plan.errors:
        status = "BLOCKED"
        error = "Input/reference errors exist; no database rows were inserted."
        mark_plan_blocked(plan)
    elif not execute:
        status = "PREVIEW"
        error = None
    elif plan.mutation_count == 0:
        status = "NO_CHANGES"
        error = None
    else:
        write_report(
            output_path,
            plan,
            status="READY_TO_EXECUTE",
            database=database,
            host=host,
            port=port,
            excel_path=input_path,
            sheet_name=sheet_name,
            execute_requested=execute,
            assume_yes=assume_yes,
            create_missing_wps=create_missing_wps,
            create_missing_units=create_missing_units,
            started_at=started_at,
            completed_at=datetime.now().astimezone(),
        )
        if not confirm_execution(database, plan, assume_yes=assume_yes):
            status = "CANCELLED"
            error = "Exact confirmation was not provided; nothing was inserted."
            for row in plan.rows:
                if row.result == "WOULD_INSERT":
                    row.result = "CANCELLED"
                    row.message = _message_with_reference_notes(
                        row, "Import was cancelled before database mutation."
                    )
            for unit in plan.units:
                if unit.result == "WOULD_INSERT":
                    unit.result = "CANCELLED"
            for wp in plan.wps:
                if wp.result == "WOULD_INSERT":
                    wp.result = "CANCELLED"
        else:
            preview_signature = plan_signature(plan)
            try:
                execution_rows = read_excel_rows(
                    input_path, sheet_name, hierarchy_id=hierarchy_id
                )
                alias = getattr(db_connection, "alias", "default")
                with transaction.atomic(using=alias):
                    with db_connection.cursor() as cursor:
                        execution_plan = build_plan(
                            cursor,
                            execution_rows,
                            hierarchy_id=hierarchy_id,
                            create_missing_wps=create_missing_wps,
                            create_missing_units=create_missing_units,
                            lock_ids=True,
                        )
                        if execution_plan.errors:
                            raise RuntimeError(
                                "Execution re-plan contains validation errors."
                            )
                        if (
                            not assume_yes
                            and plan_signature(execution_plan) != preview_signature
                        ):
                            raise RuntimeError(
                                "The planned IDs or rows changed after preview; "
                                "the whole transaction was rolled back. Run the "
                                "preview again before importing."
                            )
                        inserted_count = execute_plan(cursor, execution_plan)
                plan = execution_plan
            except Exception as exc:
                status = "ROLLED_BACK"
                error = f"{type(exc).__name__}: {exc}"
                mark_plan_blocked(plan)
                for row in plan.rows:
                    if row.result == "BLOCKED":
                        row.result = "ROLLED_BACK"
                        row.message = _message_with_reference_notes(
                            row,
                            "Transaction failed; all database changes were "
                            "rolled back.",
                        )
                for unit in plan.units:
                    if unit.result == "BLOCKED":
                        unit.result = "ROLLED_BACK"
                for wp in plan.wps:
                    if wp.result == "BLOCKED":
                        wp.result = "ROLLED_BACK"
            else:
                status = "COMMITTED"
                error = None
                if inserted_count != sum(row.result == "INSERTED" for row in plan.rows):
                    raise RuntimeError(
                        "Internal count mismatch after committed transaction."
                    )

    completed_at = datetime.now().astimezone()
    write_report(
        output_path,
        plan,
        status=status,
        database=database,
        host=host,
        port=port,
        excel_path=input_path,
        sheet_name=sheet_name,
        execute_requested=execute,
        assume_yes=assume_yes,
        create_missing_wps=create_missing_wps,
        create_missing_units=create_missing_units,
        started_at=started_at,
        completed_at=completed_at,
        error=error,
    )
    inserted_count = sum(row.result == "INSERTED" for row in plan.rows)
    print(f"Run status: {status}")
    print(f"Excel report: {output_path}")
    print(f"se_parameter rows inserted: {inserted_count}")

    result = RunResult(status, output_path, inserted_count, plan)
    if status == "ROLLED_BACK":
        raise ToolRunError(error or "Database import failed.", output_path)
    return result


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Preview or import SE master rows from Excel using Django's database. "
            "Every run writes an Excel report."
        )
    )
    parser.add_argument("--excel", type=Path, default=RUN_EXCEL_PATH)
    parser.add_argument("--sheet", default=RUN_SHEET_NAME)
    parser.add_argument("--hierarchy-id", type=int, default=RUN_HIERARCHY_ID)
    parser.add_argument("--report", type=Path, default=RUN_REPORT_PATH)
    parser.add_argument(
        "--no-create-missing-wps",
        action="store_true",
        help="Treat missing WPs as errors instead of inserting them.",
    )
    parser.add_argument(
        "--no-create-missing-units",
        action="store_true",
        help="Treat missing units as errors instead of inserting them.",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Insert rows. Without this flag the tool is preview-only.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Bypass exact interactive confirmation (audited opt-in).",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES
    create_missing_wps = RUN_CREATE_MISSING_WPS and not args.no_create_missing_wps
    create_missing_units = RUN_CREATE_MISSING_UNITS and not args.no_create_missing_units
    try:
        result = import_se_master(
            args.excel,
            sheet_name=args.sheet,
            hierarchy_id=args.hierarchy_id,
            create_missing_wps=create_missing_wps,
            create_missing_units=create_missing_units,
            execute=execute,
            assume_yes=assume_yes,
            report_path=args.report,
        )
    except ToolRunError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        print(f"Excel report: {exc.report_path}", file=sys.stderr)
        return 1
    return 1 if result.status == "BLOCKED" else 0


if __name__ == "__main__":
    raise SystemExit(main())
