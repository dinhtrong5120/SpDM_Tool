"""Preview or insert explicitly listed employees into ``user_table``.

For PyCharm's Run button, edit the RUN_* constants below. The default is a
dry run: no row is inserted unless RUN_EXECUTE is explicitly set to True.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable, Sequence

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if __name__ == "__main__":
    import os

    import django

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(0, str(PROJECT_ROOT))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

from django.db import connection, transaction  # noqa: E402

EmployeeRow = tuple[str, str, str, str]

# PyCharm Run-button configuration. Always preview before enabling execution.
RUN_EMPLOYEES: list[EmployeeRow] = [
    ("N912388", "UA3", "TSUYOSHI", "AOYAGI"),
    ("00000001", "XH4", "", ""),
    ("00000002", "XH4", "大神", ""),
    ("00000003", "XH4", "井上", ""),
    ("00000004", "XJS", "希代", ""),
    ("00000005", "XTH", "新井", ""),
    ("00000006", "", "荒舩", ""),
    ("00000007", "", "", ""),
]
RUN_EXECUTE = False
RUN_ASSUME_YES = False


def normalize_employees(rows: Iterable[Sequence[object]]) -> list[EmployeeRow]:
    """Validate rows and reject duplicate employee numbers."""
    normalized: list[EmployeeRow] = []
    seen: set[str] = set()

    for row_number, row in enumerate(rows, start=1):
        if len(row) != 4:
            raise ValueError(
                f"Row {row_number} must contain exactly four values: "
                "employee_number, section_code, first_name, last_name."
            )

        # Tất cả phải là string, nhưng chỉ employee_number bắt buộc không rỗng
        if not all(isinstance(value, str) for value in row):
            raise ValueError(
                f"Every value in row {row_number} must be a string."
            )

        employee_number = row[0].strip()
        section_code = row[1].strip()
        first_name = row[2].strip()
        last_name = row[3].strip()

        if not employee_number:
            raise ValueError(
                f"employee_number in row {row_number} must not be empty."
            )

        employee = (
            employee_number,
            section_code,
            first_name,
            last_name,
        )

        if employee_number in seen:
            raise ValueError(
                f"Duplicate employee_number: {employee_number!r}"
            )

        seen.add(employee_number)
        normalized.append(employee)

    if not normalized:
        raise ValueError("The employee list is empty.")

    return normalized


def make_placeholders(values: Sequence[object]) -> str:
    if not values:
        raise ValueError("Cannot create SQL placeholders for an empty sequence.")
    return ", ".join(["%s"] * len(values))


def database_identity() -> tuple[str, str]:
    settings_dict = connection.settings_dict
    return str(settings_dict.get("NAME") or "<unknown>"), str(
        settings_dict.get("HOST") or "localhost"
    )


def find_existing_employee_numbers(
    cursor, employee_numbers: Sequence[str]
) -> list[str]:
    placeholders = make_placeholders(employee_numbers)
    cursor.execute(
        f"""
        SELECT employee_number
        FROM user_table
        WHERE employee_number IN ({placeholders})
        ORDER BY employee_number
        """,
        employee_numbers,
    )
    return [str(row[0]) for row in cursor.fetchall()]


def print_preview(
    rows: Sequence[EmployeeRow],
    *,
    existing_employee_numbers: Sequence[str],
    database: str,
    host: str,
) -> None:
    existing = set(existing_employee_numbers)
    insert_count = sum(row[0] not in existing for row in rows)
    print("Database insert preview")
    print(f"Database: {database}")
    print(f"Host: {host}")
    print("Table: user_table")
    print(f"Rows to insert: {insert_count}")
    print(f"Rows to skip (already exist): {len(rows) - insert_count}")
    print("Selected records:")
    for employee_number, section_code, first_name, last_name in rows:
        status = "SKIP" if employee_number in existing else "INSERT"
        print(
            f"  [{status}] {employee_number} | {section_code} | "
            f"{first_name} | {last_name}"
        )


def require_confirmation(database: str, employee_numbers: Sequence[str]) -> None:
    expected = f"INSERT {database} {' '.join(employee_numbers)}"
    print("\nMutation requested. Type the following text exactly to continue:")
    print(expected)
    if input("> ").strip() != expected:
        raise RuntimeError("Confirmation did not match; nothing was inserted.")


def insert_employees(
    rows: Iterable[Sequence[object]], *, execute: bool, assume_yes: bool
) -> int:
    employees = normalize_employees(rows)
    employee_numbers = [row[0] for row in employees]
    database, host = database_identity()

    with connection.cursor() as cursor:
        existing = find_existing_employee_numbers(cursor, employee_numbers)

    print_preview(
        employees,
        existing_employee_numbers=existing,
        database=database,
        host=host,
    )
    existing_set = set(existing)
    pending = [row for row in employees if row[0] not in existing_set]

    if not pending:
        print("\nNothing to insert: all employee_number values already exist.")
        return 0

    if not execute:
        print(
            f"\nDRY RUN: {len(pending)} row(s) would be inserted; "
            "no data was changed. Use --execute to insert."
        )
        return 0

    if not assume_yes:
        require_confirmation(database, [row[0] for row in pending])

    with transaction.atomic():
        with connection.cursor() as cursor:
            # Recheck inside the transaction to reduce the duplicate-insert race.
            pending_numbers = [row[0] for row in pending]
            newly_existing = set(
                find_existing_employee_numbers(cursor, pending_numbers)
            )
            if newly_existing:
                print(
                    "Skipping employee_number values created concurrently: "
                    + ", ".join(sorted(newly_existing))
                )
                pending = [
                    row for row in pending if row[0] not in newly_existing
                ]

            if not pending:
                print("\nNothing to insert after the final duplicate check.")
                return 0

            cursor.executemany(
                """
                INSERT INTO user_table (
                    employee_number,
                    section_code,
                    contac_person_first_name,
                    contac_person_last_name
                )
                VALUES (%s, %s, %s, %s)
                """,
                pending,
            )
            if cursor.rowcount != len(pending):
                raise RuntimeError(
                    f"Unexpected inserted row count: {cursor.rowcount}; "
                    f"expected {len(pending)}."
                )

            requested_numbers = [row[0] for row in employees]
            present = find_existing_employee_numbers(cursor, requested_numbers)
            if present != sorted(requested_numbers):
                raise RuntimeError(
                    "Post-insert verification failed; the transaction will roll back."
                )

    skipped_count = len(employees) - len(pending)
    print(
        f"\nInserted and verified {len(pending)} row(s) in user_table; "
        f"skipped {skipped_count} existing row(s)."
    )
    return len(pending)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Preview or insert the configured employees into user_table."
    )
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Perform the insert; without this flag the tool is a dry run.",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip interactive confirmation (intended for audited automation).",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    insert_employees(
        RUN_EMPLOYEES,
        execute=args.execute or RUN_EXECUTE,
        assume_yes=args.yes or RUN_ASSUME_YES,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
