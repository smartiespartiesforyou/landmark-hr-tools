from collections import defaultdict
from datetime import date, datetime
import re

from openpyxl import load_workbook


GRACE_MINUTES = 7
MINIMUM_LUNCH_SHIFT_HOURS = 7
DOUBLE_SHIFT_HOURS = 12


def _clean(value):
    if value is None:
        return ""
    return str(value).strip()


def _is_yes(value):
    return _clean(value).lower() == "yes"


def _as_number(value):
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(_clean(value))
    except ValueError:
        return 0.0


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for pattern in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(_clean(value), pattern).date()
        except ValueError:
            pass
    raise ValueError(f"Could not read report date: {value}")


def _header_map(row):
    return {_clean(value): index for index, value in enumerate(row) if _clean(value)}


def _find_header(sheet, required_columns):
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        columns = _header_map(row)
        if all(column in columns for column in required_columns):
            return row_number, columns
    raise ValueError(
        "This file does not contain the expected columns: "
        + ", ".join(required_columns)
    )


def _name_key(last_first_name):
    text = re.sub(r"\s+", " ", _clean(last_first_name)).upper()
    if "," not in text:
        return text
    last, first = (part.strip() for part in text.split(",", 1))
    return f"{last}, {first}"


def _display_name(last_first_name):
    key = _name_key(last_first_name)
    if "," not in key:
        return key.title()
    last, first = (part.strip() for part in key.split(",", 1))
    return f"{first.title()} {last.title()}"


def _department(work_center):
    work_center = _clean(work_center)
    if not work_center:
        return "Department Not Listed"

    special_departments = (
        "Social Services",
        "Medical Records",
        "Nurse Aide Tr",
        "Nursing Aides",
    )
    for department in special_departments:
        if work_center.lower().startswith(department.lower()):
            return "Nurse Aide Training" if department == "Nurse Aide Tr" else department

    return work_center.split()[0]


def _is_floor_nurse(work_center):
    text = _clean(work_center).lower()
    return text.startswith("nursing lpn nurse") or text.startswith("nursing reg nurse")


def _is_overnight_cna(work_center, schedule):
    if not _clean(work_center).lower().startswith("nursing aides"):
        return False
    compact_schedule = re.sub(r"\s+", "", _clean(schedule).lower())
    return "10p-6a" in compact_schedule or "10:00p-6:00a" in compact_schedule


def _minutes(hours):
    return int(round(_as_number(hours) * 60))


def read_exception_report(path):
    sheet = load_workbook(path, data_only=True, read_only=True).active
    header_row, columns = _find_header(
        sheet,
        [
            "Last, First Name",
            "Work Center",
            "Date",
            "Work Schedule",
            "Late In",
            "Early Out",
            "Worked",
            "Not Sched.",
            "Miss Punch",
        ],
    )

    grouped = defaultdict(list)
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        name = _clean(row[columns["Last, First Name"]])
        if not name or name.lower() == "total":
            continue

        raw_date = row[columns["Date"]]
        if raw_date in (None, ""):
            continue

        work_date = _as_date(raw_date)
        key = (_name_key(name), work_date)
        grouped[key].append(
            {
                "name": name,
                "work_center": _clean(row[columns["Work Center"]]),
                "schedule": _clean(row[columns["Work Schedule"]]),
                "late_minutes": _minutes(row[columns["Late In"]]),
                "early_minutes": _minutes(row[columns["Early Out"]]),
                "worked": _as_number(row[columns["Worked"]]),
                "not_scheduled": _is_yes(row[columns["Not Sched."]]),
                "missing_punch": _as_number(row[columns["Miss Punch"]]) > 0
                or _is_yes(row[columns["Miss Punch"]]),
            }
        )

    if not grouped:
        raise ValueError("No employee rows were found in the Exception Report.")
    return grouped


def read_lunch_report(path):
    sheet = load_workbook(path, data_only=True, read_only=True).active
    header_row, columns = _find_header(
        sheet,
        ["First Name", "Last Name", "Date", "Type", "Unpaid Hours", "Paid Hours"],
    )

    lunches = defaultdict(list)
    dates = set()
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        first = _clean(row[columns["First Name"]])
        last = _clean(row[columns["Last Name"]])
        raw_date = row[columns["Date"]]
        if not first or not last or raw_date in (None, ""):
            continue

        work_date = _as_date(raw_date)
        dates.add(work_date)
        key = (_name_key(f"{last}, {first}"), work_date)
        unpaid = _as_number(row[columns["Unpaid Hours"]])
        paid = _as_number(row[columns["Paid Hours"]])
        if unpaid > 0 or paid > 0:
            lunches[key].append(
                {
                    "type": _clean(row[columns["Type"]]),
                    "hours": unpaid + paid,
                }
            )

    return lunches, dates


def build_daily_review(exception_path, lunch_path):
    attendance = read_exception_report(exception_path)
    lunches, lunch_dates = read_lunch_report(lunch_path)
    attendance_dates = {work_date for _, work_date in attendance}

    if len(attendance_dates) != 1:
        raise ValueError("The Exception Report must cover exactly one day.")
    report_date = next(iter(attendance_dates))
    if lunch_dates and lunch_dates != {report_date}:
        raise ValueError("The two reports are not for the same date.")

    missing_punches = []
    results = []

    for (name_key, work_date), rows in attendance.items():
        work_center = next((row["work_center"] for row in rows if row["work_center"]), "")
        schedule = next(
            (
                row["schedule"]
                for row in rows
                if row["schedule"] and not row["not_scheduled"]
            ),
            "",
        )
        display_name = _display_name(rows[0]["name"])

        if any(row["missing_punch"] for row in rows):
            missing_punches.append(display_name)
            continue

        scheduled_rows = [row for row in rows if not row["not_scheduled"]]
        late_minutes = max((row["late_minutes"] for row in scheduled_rows), default=0)
        early_minutes = max((row["early_minutes"] for row in scheduled_rows), default=0)
        late_minutes = late_minutes if late_minutes > GRACE_MINUTES else 0
        early_minutes = early_minutes if early_minutes > GRACE_MINUTES else 0

        worked_hours = sum(row["worked"] for row in rows)
        lunch_exempt = _is_floor_nurse(work_center) or _is_overnight_cna(
            work_center, schedule
        )
        lunch_required = worked_hours >= MINIMUM_LUNCH_SHIFT_HOURS and not lunch_exempt
        expected_breaks = 2 if worked_hours >= DOUBLE_SHIFT_HOURS else 1
        worked_segments = sum(1 for row in rows if row["worked"] > 0)
        split_breaks = max(0, worked_segments - 1)
        taken_breaks = len(lunches.get((name_key, work_date), [])) + split_breaks
        missed_required_lunch = lunch_required and taken_breaks < expected_breaks

        if not (late_minutes or early_minutes or missed_required_lunch):
            continue

        if lunch_exempt:
            lunch_status = "N/A"
        elif not lunch_required:
            lunch_status = "N/A"
        elif missed_required_lunch:
            if expected_breaks == 2 and taken_breaks == 1:
                lunch_status = "1 of 2"
            else:
                lunch_status = "No"
        else:
            lunch_status = "Yes"

        results.append(
            {
                "department": _department(work_center),
                "employee": display_name,
                "date": work_date,
                "late_minutes": late_minutes,
                "early_minutes": early_minutes,
                "lunch": lunch_status,
            }
        )

    results.sort(key=lambda row: (row["department"], row["employee"]))
    return report_date, results, sorted(set(missing_punches))
