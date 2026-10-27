"""Copy R-List responsible_employee values between variations.

Set the RUN_* constants below and run this file from PyCharm.  The default is
preview-only; set RUN_EXECUTE=True only after checking the preview.
"""

from __future__ import annotations

import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Iterable

PROJECT_ROOT = Path(__file__).resolve().parents[1]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.conf import settings  # noqa: E402
from django.db import connection, transaction  # noqa: E402

# PyCharm Run configuration. Keep RUN_EXECUTE=False for the first run.
RUN_SOURCE_PROJECT_ID = 1053
RUN_TARGET_PROJECT_ID = 1054
RUN_PHASE_ID = 16
RUN_SOURCE_VARIATION_ID = 1000012
RUN_TARGET_VARIATION_IDS = [1000014, 1000015]
RUN_EXECUTE = True
RUN_ASSUME_YES = True
REPORT_DIR = PROJECT_ROOT / "utils" / "output"

ReportRow = tuple[int, int, str | None, str | None, str | None, str]


def normalize_positive_id(raw_id: object, label: str) -> int:
    if isinstance(raw_id, bool):
        raise ValueError(f"{label} must be a positive integer, not boolean.")
    try:
        value = int(raw_id)  # type: ignore[arg-type]
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid {label}: {raw_id!r}.") from exc
    if value <= 0:
        raise ValueError(f"{label} must be greater than zero: {value}.")
    return value


def normalize_target_ids(
    raw_ids: Iterable[object], source_variation_id: int
) -> list[int]:
    target_ids: list[int] = []
    for raw_id in raw_ids:
        target_id = normalize_positive_id(raw_id, "target variation_id")
        if target_id == source_variation_id:
            raise ValueError("Source variation_id cannot also be a target.")
        if target_id not in target_ids:
            target_ids.append(target_id)
    if not target_ids:
        raise ValueError("RUN_TARGET_VARIATION_IDS cannot be empty.")
    return target_ids


def fetch_scope(
    cursor, project_id: int, phase_id: int, variation_id: int, *, lock: bool
):
    lock_sql = " FOR UPDATE" if lock else ""
    cursor.execute(
        """
        SELECT r_parameter_id, responsible_employee
        FROM project_r_parameter
        WHERE project_id = %s AND phase_id = %s AND variation_id = %s
        ORDER BY r_parameter_id
        """
        + lock_sql,
        [project_id, phase_id, variation_id],
    )
    return dict(cursor.fetchall())


def build_plan(
    cursor,
    source_project_id: int,
    target_project_id: int,
    phase_id: int,
    source_id: int,
    target_ids: list[int],
    *,
    lock: bool,
):
    source = fetch_scope(cursor, source_project_id, phase_id, source_id, lock=lock)
    if not source:
        raise ValueError(
            "No source R-List rows found for "
            f"project_id={source_project_id}, phase_id={phase_id}, "
            f"variation_id={source_id}."
        )

    plan: list[tuple[int, int, str | None, str | None]] = []
    report_rows: list[ReportRow] = []
    for target_id in target_ids:
        target = fetch_scope(cursor, target_project_id, phase_id, target_id, lock=lock)
        if not target:
            raise ValueError(
                "No target R-List rows found for "
                f"project_id={target_project_id}, phase_id={phase_id}, "
                f"variation_id={target_id}."
            )
        for r_parameter_id, source_assignee in source.items():
            if r_parameter_id not in target:
                report_rows.append(
                    (
                        target_id,
                        r_parameter_id,
                        source_assignee,
                        None,
                        None,
                        "SKIPPED_NOT_IN_TARGET",
                    )
                )
                continue
            target_assignee = target[r_parameter_id]
            action = "UNCHANGED" if target_assignee == source_assignee else "UPDATE"
            report_rows.append(
                (
                    target_id,
                    r_parameter_id,
                    source_assignee,
                    target_assignee,
                    source_assignee,
                    action,
                )
            )
            if action == "UPDATE":
                plan.append(
                    (
                        target_id,
                        r_parameter_id,
                        target_assignee,
                        source_assignee,
                    )
                )
        for r_parameter_id in sorted(set(target) - set(source)):
            target_assignee = target[r_parameter_id]
            report_rows.append(
                (
                    target_id,
                    r_parameter_id,
                    None,
                    target_assignee,
                    target_assignee,
                    "TARGET_ONLY_UNCHANGED",
                )
            )
    return plan, len(source), report_rows


def fetch_parameter_details(cursor, report_rows: list[ReportRow]):
    parameter_ids = sorted({row[1] for row in report_rows})
    placeholders = ", ".join(["%s"] * len(parameter_ids))
    cursor.execute(
        f"""
        SELECT rp.id, wp.wp, rp.performance, rp.design_item_1,
               rp.design_item_2, rp.design_item_3, u.unit
        FROM r_parameter AS rp
        LEFT JOIN wp ON wp.id = rp.wp_id
        LEFT JOIN units AS u ON u.id = rp.unit_id
        WHERE rp.id IN ({placeholders})
        """,
        parameter_ids,
    )
    return {row[0]: row[1:] for row in cursor.fetchall()}


def fetch_labels(
    cursor,
    source_project_id: int,
    target_project_id: int,
    phase_id: int,
    variation_ids: list[int],
):
    cursor.execute(
        "SELECT id, project_code FROM project_info WHERE id IN (%s, %s)",
        [source_project_id, target_project_id],
    )
    project_codes = dict(cursor.fetchall())
    missing_projects = [
        project_id
        for project_id in [source_project_id, target_project_id]
        if project_id not in project_codes
    ]
    if missing_projects:
        raise ValueError(f"project_id does not exist: {missing_projects}.")

    cursor.execute("SELECT phase FROM phase WHERE id = %s", [phase_id])
    phase_row = cursor.fetchone()
    if phase_row is None:
        raise ValueError(f"phase_id={phase_id} does not exist.")

    placeholders = ", ".join(["%s"] * len(variation_ids))
    cursor.execute(
        f"SELECT id, variation FROM variation WHERE id IN ({placeholders})",
        variation_ids,
    )
    variation_names = dict(cursor.fetchall())
    missing = [
        variation_id
        for variation_id in variation_ids
        if variation_id not in variation_names
    ]
    if missing:
        raise ValueError(f"variation_id does not exist: {missing}.")
    return project_codes, phase_row[0], variation_names


def print_preview(
    database_name: str,
    database_host: str,
    source_project_id: int,
    target_project_id: int,
    project_codes: dict[int, str],
    phase_id: int,
    phase_name: str,
    source_id: int,
    target_ids: list[int],
    variation_names: dict[int, str],
    source_count: int,
    plan: list[tuple[int, int, str | None, str | None]],
    report_rows: list[ReportRow],
) -> None:
    print(f"Database : {database_name} @ {database_host}")
    print(f"Phase    : {phase_id} ({phase_name})")
    print(
        f"Source   : project={source_project_id} "
        f"({project_codes[source_project_id]}), variation={source_id} "
        f"({variation_names[source_id]})"
    )
    print(
        f"Targets  : project={target_project_id} "
        f"({project_codes[target_project_id]}), variations="
        + ", ".join(f"{item} ({variation_names[item]})" for item in target_ids)
    )
    print(f"R rows   : {source_count} per variation")
    print(f"Changes  : {len(plan)} row(s)")
    print(f"Actions  : {dict(Counter(row[5] for row in report_rows))}")
    for target_id, r_parameter_id, old_assignee, new_assignee in plan:
        print(
            f"  variation={target_id}, r_parameter={r_parameter_id}: "
            f"{old_assignee!r} -> {new_assignee!r}"
        )


def write_report(
    path: Path,
    *,
    status: str,
    database_name: str,
    database_host: str,
    source_project_id: int,
    target_project_id: int,
    project_codes: dict[int, str],
    phase_id: int,
    phase_name: str,
    source_id: int,
    target_ids: list[int],
    variation_names: dict[int, str],
    report_rows: list[ReportRow],
    parameter_details: dict[int, tuple],
    error: str | None = None,
) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "R-List responsible employee"
    sheet.sheet_view.showGridLines = False

    sheet.merge_cells("A2:P2")
    sheet["A2"] = "R-List responsible_employee copy report"
    sheet["A2"].font = Font(name="Arial", size=14, bold=True)

    metadata = [
        ("Status", status),
        ("Generated at", datetime.now().astimezone().replace(tzinfo=None)),
        ("Database", database_name),
        ("Host", database_host),
        (
            "Source",
            f"project={source_project_id} ({project_codes[source_project_id]}), "
            f"variation={source_id} ({variation_names[source_id]})",
        ),
        (
            "Targets",
            f"project={target_project_id} ({project_codes[target_project_id]}), "
            + ", ".join(
                f"variation={item} ({variation_names[item]})" for item in target_ids
            ),
        ),
        ("Phase", f"{phase_id} ({phase_name})"),
        ("UPDATE rows", sum(row[5] == "UPDATE" for row in report_rows)),
        (
            "Skipped source-only rows",
            sum(row[5] == "SKIPPED_NOT_IN_TARGET" for row in report_rows),
        ),
    ]
    if error:
        metadata.append(("Error", error))
    for row_index, (label, value) in enumerate(metadata, start=4):
        sheet.cell(row_index, 1, label).font = Font(name="Arial", bold=True)
        sheet.cell(row_index, 2, value)
        sheet.cell(row_index, 2).alignment = Alignment(
            horizontal="left", vertical="top"
        )
        if isinstance(value, datetime):
            sheet.cell(row_index, 2).number_format = "yyyy-mm-dd hh:mm:ss"
        sheet.merge_cells(
            start_row=row_index,
            start_column=2,
            end_row=row_index,
            end_column=16,
        )

    header_row = len(metadata) + 5
    headers = [
        "Source project ID",
        "Source variation ID",
        "Target project ID",
        "Target variation ID",
        "Phase ID",
        "R parameter ID",
        "WP",
        "Performance",
        "Design item 1",
        "Design item 2",
        "Design item 3",
        "Unit",
        "Source responsible_employee",
        "Target responsible_employee before",
        "Target responsible_employee after",
        "Action",
    ]
    for column_index, header in enumerate(headers, start=1):
        sheet.cell(header_row, column_index, header)

    for (
        target_id,
        r_parameter_id,
        source_assignee,
        old_assignee,
        new_assignee,
        action,
    ) in report_rows:
        wp_name, performance, design_1, design_2, design_3, unit = (
            parameter_details.get(r_parameter_id, (None, None, None, None, None, None))
        )
        sheet.append(
            [
                source_project_id,
                source_id,
                target_project_id,
                target_id,
                phase_id,
                r_parameter_id,
                wp_name,
                performance,
                design_1,
                design_2,
                design_3,
                unit,
                source_assignee,
                old_assignee,
                new_assignee,
                action,
            ]
        )

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(name="Arial", size=10, color="FFFFFF", bold=True)
    for cell in sheet[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    action_fills = {
        "UPDATE": "C6EFCE",
        "UNCHANGED": "E7E6E6",
        "SKIPPED_NOT_IN_TARGET": "FFF2CC",
        "TARGET_ONLY_UNCHANGED": "D9EAF7",
    }
    for row in sheet.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(vertical="top")
        row[-1].fill = PatternFill("solid", fgColor=action_fills[row[-1].value])

    sheet.freeze_panes = f"A{header_row + 1}"
    sheet.auto_filter.ref = f"A{header_row}:P{sheet.max_row}"
    sheet.print_title_rows = f"{header_row}:{header_row}"
    sheet.print_area = f"A1:P{sheet.max_row}"
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    widths = [18, 19, 18, 19, 10, 16, 16, 18, 28, 28, 34, 12, 19, 23, 22, 26]
    for column_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width

    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def copy_rlist_assignees(
    source_project_id: object,
    target_project_id: object,
    phase_id: object,
    source_variation_id: object,
    target_variation_ids: Iterable[object],
    *,
    execute: bool = False,
    assume_yes: bool = False,
) -> int:
    source_project_id = normalize_positive_id(source_project_id, "source project_id")
    target_project_id = normalize_positive_id(target_project_id, "target project_id")
    phase_id = normalize_positive_id(phase_id, "phase_id")
    source_id = normalize_positive_id(source_variation_id, "source variation_id")
    target_ids = normalize_target_ids(target_variation_ids, source_id)
    all_variation_ids = [source_id, *target_ids]

    database_host = str(settings.DATABASES["default"].get("HOST") or "localhost")
    with connection.cursor() as cursor:
        cursor.execute("SELECT DATABASE()")
        database_name = str(cursor.fetchone()[0])
        project_codes, phase_name, variation_names = fetch_labels(
            cursor,
            source_project_id,
            target_project_id,
            phase_id,
            all_variation_ids,
        )
        plan, source_count, report_rows = build_plan(
            cursor,
            source_project_id,
            target_project_id,
            phase_id,
            source_id,
            target_ids,
            lock=False,
        )
        parameter_details = fetch_parameter_details(cursor, report_rows)

    report_path = REPORT_DIR / (
        f"copy_rlist_responsible_employee_{source_project_id}_to_"
        f"{target_project_id}_"
        f"{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx"
    )

    def save_report(status: str, error: str | None = None) -> None:
        write_report(
            report_path,
            status=status,
            database_name=database_name,
            database_host=database_host,
            source_project_id=source_project_id,
            target_project_id=target_project_id,
            project_codes=project_codes,
            phase_id=phase_id,
            phase_name=phase_name,
            source_id=source_id,
            target_ids=target_ids,
            variation_names=variation_names,
            report_rows=report_rows,
            parameter_details=parameter_details,
            error=error,
        )
        print(f"Excel report: {report_path}")

    print_preview(
        database_name,
        database_host,
        source_project_id,
        target_project_id,
        project_codes,
        phase_id,
        phase_name,
        source_id,
        target_ids,
        variation_names,
        source_count,
        plan,
        report_rows,
    )
    if not execute:
        save_report("PREVIEW")
        print("DRY RUN: no data was changed.")
        return 0
    if not plan:
        save_report("NO_CHANGES")
        print("Nothing to update.")
        return 0

    confirmation = (
        f"COPY {database_name} source_project={source_project_id} "
        f"target_project={target_project_id} phase={phase_id} "
        f"source_variation={source_id} "
        f"target_variations={','.join(map(str, target_ids))}"
    )
    if (
        not assume_yes
        and input(f'Type exactly "{confirmation}" to continue: ').strip()
        != confirmation
    ):
        save_report("CANCELLED")
        print("Cancelled: confirmation did not match.")
        return 0

    save_report("READY_TO_EXECUTE")
    try:
        with transaction.atomic(), connection.cursor() as cursor:
            locked_plan, _, locked_report_rows = build_plan(
                cursor,
                source_project_id,
                target_project_id,
                phase_id,
                source_id,
                target_ids,
                lock=True,
            )
            if locked_plan != plan or locked_report_rows != report_rows:
                raise RuntimeError(
                    "R-List data changed after preview; rerun the script."
                )

            updated = 0
            for target_id, r_parameter_id, old_assignee, new_assignee in plan:
                cursor.execute(
                    """
                    UPDATE project_r_parameter
                    SET responsible_employee = %s
                    WHERE project_id = %s AND r_parameter_id = %s
                      AND phase_id = %s AND variation_id = %s
                      AND responsible_employee <=> %s
                    """,
                    [
                        new_assignee,
                        target_project_id,
                        r_parameter_id,
                        phase_id,
                        target_id,
                        old_assignee,
                    ],
                )
                updated += cursor.rowcount
            if updated != len(plan):
                raise RuntimeError(
                    f"Expected to update {len(plan)} rows, but updated {updated}; "
                    "rolled back."
                )
    except Exception as exc:
        save_report("ROLLED_BACK", str(exc))
        raise

    save_report("COMMITTED")
    print(f"Completed: copied {updated} responsible_employee value(s).")
    return updated


def main() -> None:
    try:
        copy_rlist_assignees(
            RUN_SOURCE_PROJECT_ID,
            RUN_TARGET_PROJECT_ID,
            RUN_PHASE_ID,
            RUN_SOURCE_VARIATION_ID,
            RUN_TARGET_VARIATION_IDS,
            execute=RUN_EXECUTE,
            assume_yes=RUN_ASSUME_YES,
        )
    except (RuntimeError, ValueError) as exc:
        raise SystemExit(f"ERROR: {exc}") from exc


if __name__ == "__main__":
    main()
