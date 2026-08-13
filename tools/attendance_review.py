from collections import defaultdict
from datetime import date, datetime
import re

from openpyxl import load_workbook


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


def _find_header_aliases(sheet, required_aliases):
    for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        raw_columns = _header_map(row)
        normalized_columns = {
            re.sub(r"^[^:]+:\s*", "", heading).strip(): index
            for heading, index in raw_columns.items()
        }
        columns = {}
        for field, aliases in required_aliases.items():
            match = next(
                (
                    normalized_columns[alias]
                    for alias in aliases
                    if alias in normalized_columns
                ),
                None,
            )
            if match is None:
                break
            columns[field] = match
        if len(columns) == len(required_aliases):
            return row_number, columns

    raise ValueError(
        "The Lunch & Breaks Report headings were not recognized. "
        "Please download the UKG Lunch & Breaks Report without renaming its columns."
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

    if "/" in work_center:
        parts = work_center.upper().split("/")
        code = parts[3] if len(parts) > 3 else ""
        return {
            "ACT": "Activities",
            "ADM": "Administrative",
            "DTY": "Dietary",
            "HSK": "Housekeeping",
            "LDY": "Laundry",
            "MRD": "Medical Records",
            "MTN": "Maintenance",
            "NAT": "Nurse Aide Training",
            "NSG": "Nursing Aides" if parts[-1] == "CNA" else "Nursing",
            "SSV": "Social Services",
        }.get(code, "Department Not Listed")

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
    return (
        text.startswith("nursing lpn nurse")
        or text.startswith("nursing reg nurse")
        or text.endswith("/nsg/lpn")
        or text.endswith("/nsg/rgn")
    )


def _is_overnight_cna(work_center, schedule):
    center = _clean(work_center).lower()
    if not (center.startswith("nursing aides") or center.endswith("/nsg/cna")):
        return False
    compact_schedule = re.sub(r"\s+", "", _clean(schedule).lower())
    return "10p-6a" in compact_schedule or "10:00p-6:00a" in compact_schedule


def _minutes(hours):
    return int(round(_as_number(hours) * 60))


def _time_minutes(value):
    text = re.sub(r"\s+", "", _clean(value).lower())
    if not text or text == "-":
        return None
    match = re.fullmatch(r"(\d{1,2}):(\d{2})([ap])m?", text)
    if not match:
        return None
    hour, minute, period = match.groups()
    hour = int(hour) % 12
    if period == "p":
        hour += 12
    return hour * 60 + int(minute)


def _align_time(minutes, anchor):
    if minutes is None or anchor is None:
        return minutes
    return min((minutes - 1440, minutes, minutes + 1440), key=lambda x: abs(x - anchor))


def _format_clock(minutes):
    if minutes is None:
        return "-"
    minutes %= 1440
    hour24, minute = divmod(minutes, 60)
    period = "AM" if hour24 < 12 else "PM"
    hour = hour24 % 12 or 12
    return f"{hour}:{minute:02d} {period}"


def read_exception_report(path):
    sheet = load_workbook(path, data_only=True, read_only=True).active
    header_row, columns = _find_header(
        sheet,
        [
            "First Name",
            "Last Name",
            "Actual Location Full Path",
            "Date",
            "Sch. Time In",
            "Sch. Time Out",
            "Actual Time In",
            "Actual Time Out",
            "Work Hours",
            "# Incomplete Time Entries",
            "Scheduled But Absent",
            "Not Scheduled But Worked",
        ],
    )

    grouped = defaultdict(list)
    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        first = _clean(row[columns["First Name"]])
        last = _clean(row[columns["Last Name"]])
        if not first or not last:
            continue

        raw_date = row[columns["Date"]]
        if raw_date in (None, ""):
            continue

        work_date = _as_date(raw_date)
        name = f"{last}, {first}"
        key = (_name_key(name), work_date)
        scheduled_in = _time_minutes(row[columns["Sch. Time In"]])
        scheduled_out = _time_minutes(row[columns["Sch. Time Out"]])
        if scheduled_in is not None and scheduled_out is not None and scheduled_out <= scheduled_in:
            scheduled_out += 1440
        actual_in = _align_time(
            _time_minutes(row[columns["Actual Time In"]]), scheduled_in
        )
        actual_out = _align_time(
            _time_minutes(row[columns["Actual Time Out"]]), scheduled_out
        )
        grouped[key].append(
            {
                "name": name,
                "work_center": _clean(row[columns["Actual Location Full Path"]]),
                "schedule": (
                    f"{_clean(row[columns['Sch. Time In']])}-"
                    f"{_clean(row[columns['Sch. Time Out']])}"
                ),
                "scheduled_in": scheduled_in,
                "scheduled_out": scheduled_out,
                "actual_in": actual_in,
                "actual_out": actual_out,
                "worked": _as_number(row[columns["Work Hours"]]),
                "absent": _is_yes(row[columns["Scheduled But Absent"]]),
                "not_scheduled": _is_yes(row[columns["Not Scheduled But Worked"]]),
                "missing_punch": (
                    _as_number(row[columns["# Incomplete Time Entries"]]) > 0
                    or (
                        _as_number(row[columns["Work Hours"]]) > 0
                        and (actual_in is None or actual_out is None)
                    )
                ),
            }
        )

    if not grouped:
        raise ValueError("No employee rows were found in the Exception Report.")
    return grouped


def read_lunch_report(path):
    sheet = load_workbook(path, data_only=True, read_only=True).active
    header_row, columns = _find_header_aliases(
        sheet,
        {
            "First Name": ("First Name",),
            "Last Name": ("Last Name",),
            "Date": ("Date",),
            "Type": ("Type", "Lunch/Break Type"),
            "Unpaid Hours": ("Unpaid Hours", "Lunch/Break Unpaid Hours"),
        },
    )

    header_values = next(
        sheet.iter_rows(
            min_row=header_row, max_row=header_row, values_only=True
        )
    )
    normalized_header = {
        re.sub(r"^[^:]+:\s*", "", heading).strip(): index
        for heading, index in _header_map(header_values).items()
    }
    paid_hours_column = normalized_header.get("Paid Hours")

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
        paid = (
            _as_number(row[paid_hours_column])
            if paid_hours_column is not None
            else 0.0
        )
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

        # UKG lists absent employees in the Exception Report, but they must not
        # be evaluated for late/early or lunch compliance.
        if any(row["absent"] for row in rows):
            continue

        if any(row["missing_punch"] for row in rows):
            results.append(
                {
                    "department": "Incomplete Punches / Still Working",
                    "employee": display_name,
                    "date": work_date,
                    "clock_in": "-",
                    "clock_out": "-",
                    "lunch": "Review",
                }
            )
            continue

        scheduled_rows = [
            row
            for row in rows
            if not row["not_scheduled"]
            and row["scheduled_in"] is not None
            and row["scheduled_out"] is not None
        ]
        arrived_late = any(
            row["actual_in"] is not None and row["actual_in"] > row["scheduled_in"]
            for row in scheduled_rows
        )
        left_early = any(
            row["actual_out"] is not None and row["actual_out"] < row["scheduled_out"]
            for row in scheduled_rows
        )

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

        if not (arrived_late or left_early or missed_required_lunch):
            continue

        if lunch_exempt:
            lunch_status = "N/A"
        elif missed_required_lunch:
            if expected_breaks == 2 and taken_breaks == 1:
                lunch_status = "1 of 2"
            else:
                lunch_status = "No"
        elif taken_breaks:
            lunch_status = "Yes"
        else:
            lunch_status = "N/A"

        actual_ins = [row["actual_in"] for row in rows if row["actual_in"] is not None]
        actual_outs = [row["actual_out"] for row in rows if row["actual_out"] is not None]

        results.append(
            {
                "department": _department(work_center),
                "employee": display_name,
                "date": work_date,
                "clock_in": _format_clock(min(actual_ins) if actual_ins else None),
                "clock_out": _format_clock(max(actual_outs) if actual_outs else None),
                "lunch": lunch_status,
            }
        )

    results.sort(key=lambda row: (row["department"], row["employee"]))
    return report_date, results, []
