"""Preview or import R master data from an Excel workbook safely.

For PyCharm's Run button, edit the RUN_* constants below. The default is a
preview: the tool reads the configured database and writes an Excel report,
but does not insert database rows unless RUN_EXECUTE is explicitly True.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402


# PyCharm Run-button configuration. Always preview before enabling execution.
RUN_EXCEL_FILE = "/Users/TrongKieu/Downloads/R_LIST_MASTER_項目追加.xlsx"
RUN_SHEET_NAME: str | int = 0
RUN_OUTPUT_DIRECTORY: str | None = None
RUN_DEFAULT_HIERARCHY: str | None = "車両"
RUN_EXECUTE = False
RUN_ASSUME_YES = False

REQUIRED_COLUMNS = (
    "性能",
    "WP",
    "設計項目（大）",
    "設計項目（小）",
    "Unit",
)
DESIGN_ITEM_3_ALIASES = ("(Requitrement)", "(Requirement)", "目標性能")
REFERENCE_TABLES = {
    "hierarchy": "hierarchy",
    "wp": "wp",
    "units": "unit",
}
ID_TABLES = frozenset({*REFERENCE_TABLES, "r_parameter"})


@dataclass(frozen=True)
class ImportRow:
    excel_row: int
    hierarchy: str | None
    performance: str | None
    wp: str | None
    design_item_1: str | None
    design_item_2: str | None
    design_item_3: str | None
    unit: str | None
    index: int | None


@dataclass
class ImportResult:
    input_rows: list[dict[str, Any]] = field(default_factory=list)
    hierarchy_log: list[dict[str, Any]] = field(default_factory=list)
    wp_log: list[dict[str, Any]] = field(default_factory=list)
    units_log: list[dict[str, Any]] = field(default_factory=list)
    r_parameter_log: list[dict[str, Any]] = field(default_factory=list)
    errors: list[dict[str, Any]] = field(default_factory=list)
    inserted_ids: dict[str, set[int]] = field(
        default_factory=lambda: {
            "hierarchy": set(),
            "wp": set(),
            "units": set(),
            "r_parameter": set(),
        }
    )
    existing_r_parameter_count: int = 0
    duplicate_input_count: int = 0
    skipped_count: int = 0

    @property
    def pending_mutation_count(self) -> int:
        return sum(len(ids) for ids in self.inserted_ids.values())


class IdManager:
    """Allocate IDs for the four reviewed tables that do not auto-increment."""

    def __init__(self, cursor: Any, *, lock_rows: bool) -> None:
        self.cursor = cursor
        self.lock_rows = lock_rows
        self.next_ids: dict[str, int] = {}

    def get_next_id(self, table_name: str) -> int:
        if table_name not in ID_TABLES:
            raise ValueError(f"Table is not allowed for ID allocation: {table_name}")

        if table_name not in self.next_ids:
            lock_clause = " FOR UPDATE" if self.lock_rows else ""
            self.cursor.execute(
                f"SELECT id FROM `{table_name}` "
                f"ORDER BY id DESC LIMIT 1{lock_clause}"
            )
            row = self.cursor.fetchone()
            self.next_ids[table_name] = (int(row[0]) if row else 0) + 1

        allocated_id = self.next_ids[table_name]
        self.next_ids[table_name] += 1
        return allocated_id


def is_missing(value: Any) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def normalize_value(value: Any) -> str | None:
    if is_missing(value):
        return None
    normalized = str(value).strip()
    return normalized or None


def normalize_integer(value: Any) -> int | None:
    if is_missing(value) or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, bool):
        raise ValueError("No must be an integer, not a boolean.")
    try:
        numeric_value = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"No is not an integer: {value!r}") from exc
    if not numeric_value.is_integer():
        raise ValueError(f"No is not an integer: {value!r}")
    return int(numeric_value)


def safe_excel_value(value: Any) -> Any:
    return None if is_missing(value) else value


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create SQL placeholders for an empty sequence.")
    return ", ".join(["%s"] * len(values))


def resolve_design_item_3_column(columns: Iterable[object]) -> str:
    normalized_columns = {str(column).strip() for column in columns}
    for alias in DESIGN_ITEM_3_ALIASES:
        if alias in normalized_columns:
            return alias
    raise ValueError(
        "Excel is missing the requirement column. Accepted names: "
        + ", ".join(DESIGN_ITEM_3_ALIASES)
    )


def database_identity() -> tuple[str, str]:
    settings_dict = connection.settings_dict
    return (
        str(settings_dict.get("NAME") or "<unknown>"),
        str(settings_dict.get("HOST") or "localhost"),
    )


def load_excel_rows(
    excel_path: Path,
    sheet_name: str | int,
    *,
    default_hierarchy: str | None,
) -> tuple[list[ImportRow], ImportResult]:
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file was not found: {excel_path}")

    dataframe = pd.read_excel(excel_path, sheet_name=sheet_name, dtype=object)
    dataframe.columns = [str(column).strip() for column in dataframe.columns]

    missing_columns = [
        column for column in REQUIRED_COLUMNS if column not in dataframe.columns
    ]
    if default_hierarchy is None and "階層" not in dataframe.columns:
        missing_columns.append("階層")
    if missing_columns:
        raise ValueError(
            "Excel is missing required columns: " + ", ".join(missing_columns)
        )

    design_item_3_column = resolve_design_item_3_column(dataframe.columns)
    if default_hierarchy is not None:
        dataframe["階層"] = default_hierarchy

    result = ImportResult()
    valid_rows: list[ImportRow] = []

    for row_index, series in dataframe.reset_index(drop=True).iterrows():
        excel_row = int(row_index) + 2
        output_row = {
            column: safe_excel_value(series[column]) for column in dataframe.columns
        }
        output_row.update(
            {
                "excel_row": excel_row,
                "r_parameter_id": None,
                "hierarchy_id": None,
                "wp_id": None,
                "unit_id": None,
                "import_status": None,
                "import_note": None,
            }
        )

        raw_business_values = [
            series[column]
            for column in (*REQUIRED_COLUMNS, design_item_3_column)
        ]
        if all(normalize_value(value) is None for value in raw_business_values):
            output_row["import_status"] = "SKIPPED"
            output_row["import_note"] = "Empty row"
            result.skipped_count += 1
            result.input_rows.append(output_row)
            continue

        try:
            import_row = ImportRow(
                excel_row=excel_row,
                hierarchy=normalize_value(series.get("階層")),
                performance=normalize_value(series["性能"]),
                wp=normalize_value(series["WP"]),
                design_item_1=normalize_value(series["設計項目（大）"]),
                design_item_2=normalize_value(series["設計項目（小）"]),
                design_item_3=normalize_value(series[design_item_3_column]),
                unit=normalize_value(series["Unit"]),
                index=normalize_integer(series.get("No")),
            )
        except ValueError as exc:
            output_row["import_status"] = "ERROR"
            output_row["import_note"] = str(exc)
            result.errors.append(
                {
                    "excel_row": excel_row,
                    "error_type": type(exc).__name__,
                    "error_detail": str(exc),
                }
            )
        else:
            valid_rows.append(import_row)

        result.input_rows.append(output_row)

    return valid_rows, result


def get_or_create_reference(
    cursor: Any,
    id_manager: IdManager,
    *,
    table_name: str,
    value: str | None,
    staged_values: dict[str, dict[str, int]],
    execute: bool,
) -> tuple[int | None, bool]:
    if table_name not in REFERENCE_TABLES:
        raise ValueError(f"Unsupported reference table: {table_name}")
    if value is None:
        return None, False

    if value in staged_values[table_name]:
        return staged_values[table_name][value], False

    column_name = REFERENCE_TABLES[table_name]
    cursor.execute(
        f"SELECT id FROM `{table_name}` "
        f"WHERE TRIM(`{column_name}`) <=> %s ORDER BY id LIMIT 1",
        [value],
    )
    existing = cursor.fetchone()
    if existing is not None:
        return int(existing[0]), False

    reference_id = id_manager.get_next_id(table_name)
    staged_values[table_name][value] = reference_id
    if execute:
        if table_name == "wp":
            cursor.execute(
                "INSERT INTO wp (id, wp, summary_index) VALUES (%s, %s, NULL)",
                [reference_id, value],
            )
        else:
            cursor.execute(
                f"INSERT INTO `{table_name}` (id, `{column_name}`) VALUES (%s, %s)",
                [reference_id, value],
            )
        if cursor.rowcount != 1:
            raise RuntimeError(
                f"Unexpected insert count for {table_name}: {cursor.rowcount}"
            )
    return reference_id, True


def insert_r_parameter(
    cursor: Any,
    *,
    r_parameter_id: int,
    hierarchy_id: int | None,
    wp_id: int | None,
    unit_id: int | None,
    row: ImportRow,
) -> None:
    cursor.execute(
        """
        INSERT INTO r_parameter (
            id, performance, design_item_1, design_item_2, design_item_3,
            unit, hierarchy_id, wp_id, `index`, `condition`, unit_id
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, NULL, %s)
        """,
        [
            r_parameter_id,
            row.performance,
            row.design_item_1,
            row.design_item_2,
            row.design_item_3,
            row.unit,
            hierarchy_id,
            wp_id,
            row.index,
            unit_id,
        ],
    )
    if cursor.rowcount != 1:
        raise RuntimeError(
            f"Unexpected insert count for r_parameter: {cursor.rowcount}"
        )


def process_rows(
    cursor: Any,
    rows: Sequence[ImportRow],
    result: ImportResult,
    *,
    execute: bool,
) -> ImportResult:
    id_manager = IdManager(cursor, lock_rows=execute)
    staged_values = {table_name: {} for table_name in REFERENCE_TABLES}
    staged_parameters: dict[tuple[object, ...], int] = {}
    output_by_row = {row["excel_row"]: row for row in result.input_rows}

    for row in rows:
        hierarchy_id, new_hierarchy = get_or_create_reference(
            cursor,
            id_manager,
            table_name="hierarchy",
            value=row.hierarchy,
            staged_values=staged_values,
            execute=execute,
        )
        wp_id, new_wp = get_or_create_reference(
            cursor,
            id_manager,
            table_name="wp",
            value=row.wp,
            staged_values=staged_values,
            execute=execute,
        )
        unit_id, new_unit = get_or_create_reference(
            cursor,
            id_manager,
            table_name="units",
            value=row.unit,
            staged_values=staged_values,
            execute=execute,
        )

        reference_data = (
            (
                "hierarchy",
                hierarchy_id,
                new_hierarchy,
                {"hierarchy": row.hierarchy},
                result.hierarchy_log,
            ),
            (
                "wp",
                wp_id,
                new_wp,
                {"wp": row.wp, "summary_index": None},
                result.wp_log,
            ),
            (
                "units",
                unit_id,
                new_unit,
                {"unit": row.unit},
                result.units_log,
            ),
        )
        for table_name, reference_id, is_new, values, log in reference_data:
            if reference_id is None or not is_new:
                continue
            result.inserted_ids[table_name].add(reference_id)
            log.append(
                {
                    "excel_row": row.excel_row,
                    "id": reference_id,
                    **values,
                    "status": "INSERTED" if execute else "WOULD_INSERT",
                }
            )

        output_row = output_by_row[row.excel_row]
        output_row.update(
            {
                "hierarchy_id": hierarchy_id,
                "wp_id": wp_id,
                "unit_id": unit_id,
            }
        )

        resolved_key = (
            hierarchy_id,
            wp_id,
            row.performance,
            row.design_item_1,
            row.design_item_2,
            row.design_item_3,
            row.unit,
            unit_id,
        )
        if resolved_key in staged_parameters:
            r_parameter_id = staged_parameters[resolved_key]
            status = "DUPLICATE_INPUT"
            note = "Same business key already appeared earlier in this workbook"
            result.duplicate_input_count += 1
        else:
            cursor.execute(
                """
                SELECT MIN(id)
                FROM r_parameter
                WHERE hierarchy_id <=> %s
                  AND wp_id <=> %s
                  AND TRIM(performance) <=> %s
                  AND TRIM(design_item_1) <=> %s
                  AND TRIM(design_item_2) <=> %s
                  AND TRIM(design_item_3) <=> %s
                  AND TRIM(unit) <=> %s
                  AND unit_id <=> %s
                """,
                list(resolved_key),
            )
            existing = cursor.fetchone()
            existing_id = (
                int(existing[0]) if existing and existing[0] is not None else None
            )
            if existing_id is not None:
                r_parameter_id = existing_id
                status = "EXISTING"
                note = "Business key already exists in the database"
                result.existing_r_parameter_count += 1
            else:
                r_parameter_id = id_manager.get_next_id("r_parameter")
                staged_parameters[resolved_key] = r_parameter_id
                if execute:
                    insert_r_parameter(
                        cursor,
                        r_parameter_id=r_parameter_id,
                        hierarchy_id=hierarchy_id,
                        wp_id=wp_id,
                        unit_id=unit_id,
                        row=row,
                    )
                result.inserted_ids["r_parameter"].add(r_parameter_id)
                status = "INSERTED" if execute else "WOULD_INSERT"
                note = (
                    "Inserted in the current transaction"
                    if execute
                    else "No database row was inserted"
                )

        output_row.update(
            {
                "r_parameter_id": r_parameter_id,
                "import_status": status,
                "import_note": note,
            }
        )
        result.r_parameter_log.append(
            {
                "excel_row": row.excel_row,
                "id": r_parameter_id,
                "performance": row.performance,
                "design_item_1": row.design_item_1,
                "design_item_2": row.design_item_2,
                "design_item_3": row.design_item_3,
                "unit": row.unit,
                "hierarchy_id": hierarchy_id,
                "wp_id": wp_id,
                "index": row.index,
                "condition": None,
                "unit_id": unit_id,
                "status": status,
            }
        )

    return result


def verify_inserted_rows(cursor: Any, inserted_ids: dict[str, set[int]]) -> None:
    for table_name, ids in inserted_ids.items():
        if not ids:
            continue
        ordered_ids = sorted(ids)
        placeholders = make_placeholders(ordered_ids)
        cursor.execute(
            f"SELECT COUNT(*) FROM `{table_name}` WHERE id IN ({placeholders})",
            ordered_ids,
        )
        actual_count = int(cursor.fetchone()[0])
        if actual_count != len(ordered_ids):
            raise RuntimeError(
                f"Post-insert verification failed for {table_name}: "
                f"found {actual_count}, expected {len(ordered_ids)}."
            )


def print_preview(
    result: ImportResult,
    *,
    database: str,
    host: str,
    excel_path: Path,
) -> None:
    print("R master import preview")
    print(f"Database: {database}")
    print(f"Host: {host}")
    print(f"Input file: {excel_path.resolve()}")
    print("Affected tables and estimated new rows:")
    for table_name in ("hierarchy", "wp", "units", "r_parameter"):
        print(f"  {table_name}: {len(result.inserted_ids[table_name])}")
    print(f"Existing r_parameter rows: {result.existing_r_parameter_count}")
    print(f"Duplicate input rows: {result.duplicate_input_count}")
    print(f"Skipped empty rows: {result.skipped_count}")
    print(f"Validation errors: {len(result.errors)}")
    print("Selected records:")
    for row in result.r_parameter_log:
        print(
            f"  [{row['status']}] Excel row {row['excel_row']} -> "
            f"r_parameter.id={row['id']} | wp_id={row['wp_id']} | "
            f"{row['performance']} | {row['design_item_1']} | "
            f"{row['design_item_2']} | {row['design_item_3']} | {row['unit']}"
        )


def format_target_ids(inserted_ids: dict[str, set[int]]) -> str:
    parts = []
    for table_name in ("hierarchy", "wp", "units", "r_parameter"):
        ids = ",".join(str(value) for value in sorted(inserted_ids[table_name]))
        parts.append(f"{table_name}={ids or '-'}")
    return ";".join(parts)


def require_confirmation(
    database: str, inserted_ids: dict[str, set[int]]
) -> None:
    expected = f"IMPORT {database} {format_target_ids(inserted_ids)}"
    print("\nMutation requested. Type the following text exactly to continue:")
    print(expected)
    if input("> ").strip() != expected:
        raise RuntimeError("Confirmation did not match; nothing was inserted.")


def auto_adjust_excel_columns(writer: pd.ExcelWriter) -> None:
    for worksheet in writer.book.worksheets:
        for column_cells in worksheet.columns:
            max_length = max(
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=0,
            )
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(
                max_length + 2, 50
            )
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions


def write_report(
    output_path: Path,
    result: ImportResult,
    *,
    mode: str,
    transaction_result: str,
    database: str,
    host: str,
    excel_path: Path,
) -> None:
    summary = [
        {"item": "Mode", "value": mode},
        {"item": "Transaction result", "value": transaction_result},
        {"item": "Database", "value": database},
        {"item": "Host", "value": host},
        {"item": "Input file", "value": str(excel_path.resolve())},
        {"item": "Excel rows", "value": len(result.input_rows)},
        {"item": "New hierarchy", "value": len(result.inserted_ids["hierarchy"])},
        {"item": "New wp", "value": len(result.inserted_ids["wp"])},
        {"item": "New units", "value": len(result.inserted_ids["units"])},
        {
            "item": "New r_parameter",
            "value": len(result.inserted_ids["r_parameter"]),
        },
        {"item": "Existing r_parameter", "value": result.existing_r_parameter_count},
        {"item": "Duplicate input", "value": result.duplicate_input_count},
        {"item": "Skipped rows", "value": result.skipped_count},
        {"item": "Error rows", "value": len(result.errors)},
        {"item": "Created at", "value": datetime.now().isoformat(timespec="seconds")},
    ]
    sheets = {
        "Summary": summary,
        "Input_With_ID": sorted(result.input_rows, key=lambda row: row["excel_row"]),
        "Hierarchy_Log": result.hierarchy_log,
        "WP_Log": result.wp_log,
        "Units_Log": result.units_log,
        "R_Parameter_Log": result.r_parameter_log,
        "Errors": result.errors,
    }
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
        auto_adjust_excel_columns(writer)


def build_output_path(
    excel_path: Path, output_directory: str | None, *, execute: bool
) -> Path:
    output_dir = Path(output_directory) if output_directory else excel_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode = "IMPORT" if execute else "PREVIEW"
    return output_dir / f"{excel_path.stem}_report_{mode}_{timestamp}.xlsx"


def import_r_parameter_from_excel(
    excel_file_path: str,
    *,
    sheet_name: str | int = 0,
    output_directory: str | None = None,
    default_hierarchy: str | None = "車両",
    execute: bool = False,
    assume_yes: bool = False,
) -> str:
    excel_path = Path(excel_file_path).expanduser()
    rows, base_result = load_excel_rows(
        excel_path,
        sheet_name,
        default_hierarchy=default_hierarchy,
    )
    # Validate report-directory access before any database mutation can start.
    output_path = build_output_path(excel_path, output_directory, execute=execute)
    database, host = database_identity()

    with connection.cursor() as cursor:
        preview_result = process_rows(cursor, rows, base_result, execute=False)
    print_preview(
        preview_result,
        database=database,
        host=host,
        excel_path=excel_path,
    )

    if not execute:
        final_result = preview_result
        transaction_result = "NOT STARTED (PREVIEW ONLY)"
    else:
        # Persist the preview before confirmation. Besides leaving an audit
        # artifact when the operator cancels, this verifies that the report can
        # be written before the database transaction is allowed to start.
        write_report(
            output_path,
            preview_result,
            mode="PREVIEW BEFORE IMPORT",
            transaction_result="NOT STARTED",
            database=database,
            host=host,
            excel_path=excel_path,
        )
        if preview_result.errors:
            raise RuntimeError(
                f"Import refused because validation errors were found. Report: {output_path}"
            )

        confirmed_ids = {
            table_name: set(ids)
            for table_name, ids in preview_result.inserted_ids.items()
        }
        if preview_result.pending_mutation_count and not assume_yes:
            require_confirmation(database, confirmed_ids)

        # Re-read and re-plan inside one transaction. This prevents stale preview
        # data from causing partial inserts and locks the current maximum IDs.
        execution_rows, execution_result = load_excel_rows(
            excel_path,
            sheet_name,
            default_hierarchy=default_hierarchy,
        )
        with transaction.atomic():
            with connection.cursor() as cursor:
                final_result = process_rows(
                    cursor, execution_rows, execution_result, execute=True
                )
                if not assume_yes and final_result.inserted_ids != confirmed_ids:
                    raise RuntimeError(
                        "The target IDs changed after preview; the transaction was "
                        "rolled back. Run the preview again before importing."
                    )
                verify_inserted_rows(cursor, final_result.inserted_ids)
        transaction_result = "COMMITTED"

    write_report(
        output_path,
        final_result,
        mode="IMPORT" if execute else "PREVIEW - NO DATABASE CHANGES",
        transaction_result=transaction_result,
        database=database,
        host=host,
        excel_path=excel_path,
    )
    print(f"\nReport file: {output_path}")
    return str(output_path)


def parse_sheet_name(raw_value: str) -> str | int:
    stripped = raw_value.strip()
    try:
        return int(stripped)
    except ValueError:
        return stripped


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preview or import R master data from Excel using Django's database."
    )
    parser.add_argument("--excel", default=RUN_EXCEL_FILE, help="Input .xlsx path.")
    parser.add_argument(
        "--sheet",
        default=str(RUN_SHEET_NAME),
        help="Worksheet name or zero-based worksheet index.",
    )
    parser.add_argument(
        "--output-directory",
        default=RUN_OUTPUT_DIRECTORY,
        help="Report directory; defaults to the input workbook directory.",
    )
    parser.add_argument(
        "--default-hierarchy",
        default=RUN_DEFAULT_HIERARCHY,
        help="Hierarchy assigned to every imported row (default: 車両).",
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Insert rows; without this flag the tool only previews.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip confirmation (only for audited non-interactive runs).",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    import_r_parameter_from_excel(
        args.excel,
        sheet_name=parse_sheet_name(args.sheet),
        output_directory=args.output_directory,
        default_hierarchy=args.default_hierarchy,
        execute=args.execute or RUN_EXECUTE,
        assume_yes=args.yes or RUN_ASSUME_YES,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
