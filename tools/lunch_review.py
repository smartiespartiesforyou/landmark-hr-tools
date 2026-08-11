import re
from datetime import datetime, timedelta


def parse_time(value):
    value = value.strip().lower()

    if value.endswith("a"):
        value = value[:-1] + "AM"
    elif value.endswith("p"):
        value = value[:-1] + "PM"

    return datetime.strptime(value, "%I:%M%p")


def hours_between(start_text, end_text):
    start = parse_time(start_text)
    end = parse_time(end_text)

    if end <= start:
        end += timedelta(days=1)

    return (end - start).total_seconds() / 3600


def extract_employee_page(text):

    name_match = re.search(
        r"Employee Timesheet.*?\n(.+?)\s+\(Employee Id:",
        text,
        re.DOTALL,
    )

    department_match = re.search(
        r"Jobs \(HR\)\s+(.+)",
        text,
    )

    name = name_match.group(1).strip() if name_match else "Unknown Employee"
    department = department_match.group(1).strip() if department_match else ""

    shifts = {}
    current_date = None

    for line in text.splitlines():

        line = line.strip()

        # A timesheet can contain more than one week.  A Week Total is only a
        # subtotal, so skip it and continue reading the remaining dated rows.
        if line.startswith("Week Total:"):
            current_date = None
            continue

        # The unqualified Total row is the final total for this employee page.
        if line.startswith("Total:"):
            break

        dated_row = re.match(
            r"^[A-Z][a-z]{2}\s+"
            r"(\d{2}/\d{2}/\d{4})\s+"
            r"(\d{2}:\d{2}[ap])\s+"
            r"(?:[A-Z][a-z]{2}\s+)?"
            r"(\d{2}:\d{2}[ap])",
            line,
        )

        continuation_row = re.match(
            r"^(\d{2}:\d{2}[ap])\s+"
            r"(?:[A-Z][a-z]{2}\s+)?"
            r"(\d{2}:\d{2}[ap])",
            line,
        )

        if dated_row:

            current_date = dated_row.group(1)

            shifts.setdefault(current_date, []).append(
                (
                    dated_row.group(2),
                    dated_row.group(3),
                )
            )

        elif continuation_row and current_date:

            shifts.setdefault(current_date, []).append(
                (
                    continuation_row.group(1),
                    continuation_row.group(2),
                )
            )

    return name, department, shifts
