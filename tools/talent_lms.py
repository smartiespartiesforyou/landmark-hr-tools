import csv
import re
from calendar import monthrange
from datetime import date, datetime
from difflib import SequenceMatcher

from openpyxl import load_workbook


UKG_REQUIRED = {
    "Employee Id",
    "Last Name",
    "First Name",
    "Employee Status",
    "Date Hired",
    "Date Re-Hired",
    "DEPT",
}

TALENT_REQUIRED = {
    "User",
    "Progress status",
    "Completion date",
}


def normalize_name(value):
    value = str(value or "").casefold().strip()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None

    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def _find_header_row(rows):
    for index, row in enumerate(rows):
        values = {str(value).strip() for value in row if value is not None}
        if {"Last Name", "First Name", "Date Hired"}.issubset(values):
            return index
    raise ValueError("Could not find the employee header row in the UKG file.")


def read_ukg_employees(path, year, month):
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))

    header_row = _find_header_row(rows)
    headers = [str(value).strip() if value is not None else "" for value in rows[header_row]]
    missing = UKG_REQUIRED.difference(headers)
    if missing:
        raise ValueError("UKG file is missing: " + ", ".join(sorted(missing)))

    column = {name: index for index, name in enumerate(headers)}
    month_end = date(year, month, monthrange(year, month)[1])
    employees = []

    for row in rows[header_row + 1:]:
        employee_id = row[column["Employee Id"]]
        if employee_id in (None, ""):
            continue

        status = str(row[column["Employee Status"]] or "").strip()
        if status.casefold() in {"terminated", "deceased", "resigned", "retired"}:
            continue

        hired = _as_date(row[column["Date Hired"]])
        rehired = _as_date(row[column["Date Re-Hired"]])
        effective_hire = rehired or hired

        if not effective_hire or effective_hire > month_end:
            continue

        first = str(row[column["First Name"]] or "").strip()
        last = str(row[column["Last Name"]] or "").strip()
        department = str(row[column["DEPT"]] or "").strip()
        job = ""
        if "JOB" in column:
            job = str(row[column["JOB"]] or "").strip()

        employees.append({
            "employee_id": str(employee_id).strip(),
            "first": first,
            "last": last,
            "display_name": f"{first} {last}".strip().title(),
            "talent_key": normalize_name(f"{last} {first}"),
            "department": department,
            "job": job,
            "effective_hire": effective_hire,
        })

    workbook.close()
    return employees


def read_talent_lms(path):
    with open(path, newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        headers = set(reader.fieldnames or [])
        missing = TALENT_REQUIRED.difference(headers)
        if missing:
            raise ValueError("TalentLMS file is missing: " + ", ".join(sorted(missing)))

        records = []
        for row in reader:
            user = str(row.get("User") or "").strip()
            if not user:
                continue
            records.append({
                "user": user,
                "key": normalize_name(user),
                "status": str(row.get("Progress status") or "").strip(),
                "completion_date": _as_date(row.get("Completion date")),
            })
    return records


def _name_parts(key):
    parts = key.split()
    if not parts:
        return "", ""
    return parts[0], " ".join(parts[1:])


def find_possible_name_mismatches(employees, talent_records):
    talent_by_key = {row["key"]: row for row in talent_records}
    unmatched_talent = [row for row in talent_records if row["key"] not in {e["talent_key"] for e in employees}]
    possible = []

    for employee in employees:
        if employee["talent_key"] in talent_by_key:
            continue

        ukg_last, ukg_first = _name_parts(employee["talent_key"])
        best = None
        best_score = 0.0

        for talent in unmatched_talent:
            talent_last, talent_first = _name_parts(talent["key"])
            if ukg_last != talent_last or not ukg_first or not talent_first:
                continue

            score = SequenceMatcher(None, ukg_first, talent_first).ratio()
            prefix_match = ukg_first.startswith(talent_first) or talent_first.startswith(ukg_first)
            if (score >= 0.72 or prefix_match) and score > best_score:
                best = talent
                best_score = score

        if best:
            possible.append({
                "ukg_name": employee["display_name"],
                "talent_name": best["user"],
                "talent_status": best["status"],
            })

    return possible


def build_incomplete_list(employees, talent_records):
    talent_by_key = {}
    for record in talent_records:
        talent_by_key.setdefault(record["key"], []).append(record)

    incomplete = []
    completed = 0

    for employee in employees:
        records = talent_by_key.get(employee["talent_key"], [])
        did_complete = any(
            record["status"].casefold() == "completed" or record["completion_date"]
            for record in records
        )

        if did_complete:
            completed += 1
        else:
            incomplete.append(employee)

    incomplete.sort(key=lambda item: (item["department"], item["last"], item["first"]))
    return incomplete, completed
