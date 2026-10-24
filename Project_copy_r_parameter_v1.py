"""Copy các r_parameter_id nguồn sang nhiều project/phase đích.

Tái sử dụng utility ``Project_copy_r_parameter`` đã có để giữ một nguồn logic
duy nhất cho validation, dry-run, transaction, idempotency và báo cáo Excel.

Mỗi ``r_parameter_id`` xuất hiện trong ``project_r_parameter`` của source sẽ
được thêm vào mọi variation đang có ở từng destination. Script chỉ insert bốn
cột khóa của ``project_r_parameter``; dữ liệu nghiệp vụ khác không được copy.
"""

from __future__ import annotations

from pathlib import Path


RUN_SOURCE = (28, 27)
RUN_DESTINATIONS = [
    (1028, 27),
    (1029, 27),
    (1030, 27),
    (1031, 27),
    (1032, 27),
    (1033, 27),
    (1034, 27),
]

# Mặc định chỉ preview và xuất Excel. Chỉ bật sau khi kiểm tra báo cáo.
RUN_EXECUTE = False
RUN_ASSUME_YES = False
RUN_REPORT_DIRECTORY = Path(__file__).resolve().parent / "output"


def main() -> int:
    """Chạy bằng nút Run của PyCharm với cấu hình ở trên."""
    import os
    import sys

    import django

    project_root = Path(__file__).resolve().parents[2]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "backend.settings")
    django.setup()

    from utils.Tools_new.Project_copy_r_parameter import copy_r_parameters

    return copy_r_parameters(
        RUN_SOURCE,
        RUN_DESTINATIONS,
        execute=RUN_EXECUTE,
        assume_yes=RUN_ASSUME_YES,
        report_directory=RUN_REPORT_DIRECTORY,
    )


if __name__ == "__main__":
    raise SystemExit(main())
