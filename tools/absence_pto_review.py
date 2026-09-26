import math
import pandas as pd


def _clean(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value).strip()


def _hours(value):
    text = _clean(value)
    if not text or text == "-":
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _time(value):
    return _clean(value).replace("\n", " ").strip()


def _yes(value):
    return _clean(value).lower() in {"yes", "y", "true", "1"}


def _department(row):
    path = _clean(row.get("Actual Jobs (HR) Full Path"))
    upper = path.upper()
    job = upper.split(">")[-1].split("/")[-1].strip() if path else ""
    if "CNA" in upper:
        return "CNAs"
    if "SUNSHINE" in upper:
        return "Sunshine Aides"
    if "NSG" in upper or "NURS" in upper:
        return "Nursing"
    parts = [p.strip() for p in path.replace(">", "/").split("/") if p.strip()]
    if parts:
        for part in parts:
            if part.upper() not in {"LM SOUTH", "LANDMARK SOUTH", "STF"}:
                return part.title()
    return "Other"


def review_absence_pto(path):
    raw = pd.read_excel(path, header=None)
    header_row = None
    for index, row in raw.iterrows():
        values = [_clean(value) for value in row.tolist()]
        if "First Name" in values and "Last Name" in values and "Sch. Time In" in values:
            header_row = index
            break
    if header_row is None:
        raise ValueError("Could not find the Late / Early / Absent report headings.")

    data = pd.read_excel(path, header=header_row)
    required = {
        "First Name", "Last Name", "Date", "Sch. Time In", "Sch. Time Out",
        "Actual Time In", "Actual Time Out", "Time Off Hours", "Time Off Name",
        "Scheduled But Absent", "Actual Jobs (HR) Full Path"
    }
    missing = sorted(required.difference(data.columns))
    if missing:
        raise ValueError("This UKG report is missing: " + ", ".join(missing))

    absence_review, time_off, partial_or_review = [], [], []
    report_dates = []

    for _, row in data.iterrows():
        first, last = _clean(row.get("First Name")), _clean(row.get("Last Name"))
        if not first or not last:
            continue
        date = pd.to_datetime(row.get("Date"), errors="coerce")
        if pd.isna(date):
            continue
        report_dates.append(date.date())

        scheduled_in, scheduled_out = _time(row.get("Sch. Time In")), _time(row.get("Sch. Time Out"))
        actual_in, actual_out = _time(row.get("Actual Time In")), _time(row.get("Actual Time Out"))
        off_hours, off_name = _hours(row.get("Time Off Hours")), _clean(row.get("Time Off Name"))
        scheduled_absent = _yes(row.get("Scheduled But Absent"))
        has_any_punch = bool(actual_in or actual_out)
        has_complete_punches = bool(actual_in and actual_out)
        has_time_off = bool(off_name or off_hours)

        item = {
            "name": f"{first.title()} {last.title()}",
            "date": date.date(),
            "department": _department(row),
            "scheduled": f"{scheduled_in} – {scheduled_out}" if scheduled_in or scheduled_out else "Not scheduled",
            "actual": f"{actual_in or '—'} – {actual_out or '—'}",
            "time_off_hours": off_hours,
            "time_off_name": off_name or "Time Off",
        }

        if has_time_off:
            if has_any_punch:
                item["status"] = "Worked + Time Off"
                partial_or_review.append(item)
            else:
                item["status"] = off_name or "Time Off"
                time_off.append(item)
        elif scheduled_absent:
            item["status"] = "Absence Review"
            absence_review.append(item)
        elif (scheduled_in or scheduled_out) and not has_complete_punches and has_any_punch:
            item["status"] = "Incomplete Punch"
            partial_or_review.append(item)

    absence_review.sort(key=lambda item: (item["department"], item["name"]))
    report_date = min(report_dates) if report_dates else None
    return report_date, absence_review, time_off, partial_or_review
