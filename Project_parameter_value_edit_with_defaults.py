"""Run the project_parameter sync with opt-in defaults for blank Excel cells.

When enabled, a blank ステータス becomes approval_status=0 and a blank 更新者
becomes employee_number="00000000". Existing non-blank Excel values are never
replaced. All database safety, reporting, and optional DB_ONLY deletion behavior
comes from Project_parameter_value_edit.py.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from utils.Tools import Project_parameter_value_edit as base  # noqa: E402

# PyCharm Run-button configuration. Change these values in this file.
RUN_PROJECT_ID = base.RUN_PROJECT_ID
RUN_PHASE_ID = base.RUN_PHASE_ID
RUN_EXCEL_PATH = base.RUN_EXCEL_PATH
RUN_SHEET_NAME = base.RUN_SHEET_NAME
RUN_REPORT_PATH: Path | None = None

# Mutation and default-value switches are deliberately disabled by default.
RUN_EXECUTE = False
RUN_ASSUME_YES = False
RUN_DELETE_DB_ONLY = False
RUN_APPLY_DEFAULT_VALUES = False

DEFAULT_APPROVAL_STATUS = 0
DEFAULT_EMPLOYEE_NUMBER = "00000007"


def fill_blank_excel_defaults(
    read_results: Sequence[base.ExcelReadResult],
    *,
    enabled: bool,
) -> tuple[list[base.ExcelReadResult], int, int]:
    """Fill the two configured defaults and return per-field fill counts."""
    if not enabled:
        return list(read_results), 0, 0

    status_count = 0
    employee_count = 0
    normalized_results: list[base.ExcelReadResult] = []
    for result in read_results:
        normalized_rows: list[base.ExcelParameterRow] = []
        for row in result.rows:
            values = dict(row.values)
            if values["approval_status"] is None:
                values["approval_status"] = DEFAULT_APPROVAL_STATUS
                status_count += 1
            if values["employee_number"] is None:
                values["employee_number"] = DEFAULT_EMPLOYEE_NUMBER
                employee_count += 1
            normalized_rows.append(
                base.ExcelParameterRow(
                    excel_row=row.excel_row,
                    se_parameter_id=row.se_parameter_id,
                    values=tuple(
                        (column, values[column]) for column in base.DATA_COLUMNS
                    ),
                )
            )
        normalized_results.append(
            base.ExcelReadResult(
                variation_name=result.variation_name,
                rows=tuple(normalized_rows),
                header_row=result.header_row,
                sheet_name=result.sheet_name,
            )
        )
    return normalized_results, status_count, employee_count


def read_excel_variations_with_defaults(
    excel_path: Path | str,
    sheet_name: str,
    *,
    apply_default_values: bool,
) -> list[base.ExcelReadResult]:
    read_results = base.read_excel_variations(excel_path, sheet_name)
    normalized, status_count, employee_count = fill_blank_excel_defaults(
        read_results,
        enabled=apply_default_values,
    )
    if apply_default_values:
        print(
            "Excel defaults enabled: "
            f"approval_status={DEFAULT_APPROVAL_STATUS} applied to "
            f"{status_count} cells; employee_number="
            f"{DEFAULT_EMPLOYEE_NUMBER!r} applied to {employee_count} cells."
        )
    else:
        print(
            "Excel defaults disabled: blank approval_status and employee_number "
            "remain SQL NULL."
        )
    return normalized


def upsert_project_parameters(
    project_id: object,
    phase_id: object,
    excel_path: Path | str,
    *,
    sheet_name: str = RUN_SHEET_NAME,
    execute: bool = False,
    assume_yes: bool = False,
    delete_db_only: bool = False,
    apply_default_values: bool = False,
    report_path: Path | str | None = None,
    db_connection=base.connection,
) -> list[base.PlanRow]:
    def excel_reader(path: Path | str, requested_sheet: str):
        return read_excel_variations_with_defaults(
            path,
            requested_sheet,
            apply_default_values=apply_default_values,
        )

    return base.upsert_project_parameters(
        project_id,
        phase_id,
        excel_path,
        sheet_name=sheet_name,
        execute=execute,
        assume_yes=assume_yes,
        delete_db_only=delete_db_only,
        report_path=report_path,
        db_connection=db_connection,
        excel_reader=excel_reader,
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Sync project_parameter from Excel with optional defaults for blank "
            "approval status and employee number cells."
        )
    )
    parser.add_argument("--project-id", type=int, default=RUN_PROJECT_ID)
    parser.add_argument("--phase-id", type=int, default=RUN_PHASE_ID)
    parser.add_argument("--excel", type=Path, default=RUN_EXCEL_PATH)
    parser.add_argument("--sheet", default=RUN_SHEET_NAME)
    parser.add_argument("--report", type=Path, default=RUN_REPORT_PATH)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--delete-db-only", action="store_true")
    parser.add_argument(
        "--apply-default-values",
        action="store_true",
        help=(
            "Use approval_status=0 and employee_number='00000000' for their "
            "respective blank Excel cells."
        ),
    )
    parser.add_argument("--yes", action="store_true")
    return parser.parse_args(argv)


def main() -> None:
    args = parse_args()
    execute = args.execute or RUN_EXECUTE
    assume_yes = args.yes or RUN_ASSUME_YES
    delete_db_only = args.delete_db_only or RUN_DELETE_DB_ONLY
    apply_default_values = args.apply_default_values or RUN_APPLY_DEFAULT_VALUES
    if assume_yes and not execute:
        raise ValueError("RUN_ASSUME_YES/--yes requires execute mode.")
    upsert_project_parameters(
        args.project_id,
        args.phase_id,
        args.excel,
        sheet_name=args.sheet,
        execute=execute,
        assume_yes=assume_yes,
        delete_db_only=delete_db_only,
        apply_default_values=apply_default_values,
        report_path=args.report,
    )


if __name__ == "__main__":
    main()
