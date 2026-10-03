import os
import re
import unicodedata
from collections import defaultdict


def _normalize(value):
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = value.encode("ascii", "ignore").decode("ascii").casefold().strip()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return " ".join(value.split())


def employee_file_key(employee):
    first = str(employee.get("first") or "").strip()
    last = str(employee.get("last") or "").strip()
    if not first or not last:
        return ""
    return _normalize(f"{first[0]} {last}")


def read_pdf_name_list(path):
    names = []
    with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
        for raw_line in handle:
            filename = raw_line.strip().strip('"')
            if not filename:
                continue
            basename = os.path.basename(filename)
            if basename.casefold().endswith(".pdf"):
                basename = basename[:-4]
            key = _normalize(basename)
            if key:
                names.append({"filename": raw_line.strip(), "base": basename, "key": key})
    if not names:
        raise ValueError("No PDF names were found in PDF-List.txt. Run: dir /b *.pdf > PDF-List.txt")
    return names


def compare_signature_files(employees, pdf_names):
    employees_by_key = defaultdict(list)
    files_by_key = defaultdict(list)

    for employee in employees:
        key = employee_file_key(employee)
        if key:
            employees_by_key[key].append(employee)

    for item in pdf_names:
        files_by_key[item["key"]].append(item)

    found = []
    missing = []
    review = []
    used_file_keys = set()

    for key, employee_group in employees_by_key.items():
        file_group = files_by_key.get(key, [])
        if len(employee_group) == 1 and len(file_group) == 1:
            found.append({"employee": employee_group[0], "file": file_group[0]["filename"]})
            used_file_keys.add(key)
        elif not file_group:
            if len(employee_group) == 1:
                missing.append(employee_group[0])
            else:
                review.append({
                    "key": key,
                    "reason": "Duplicate UKG initial + last name; no matching PDF was found.",
                    "employees": employee_group,
                    "files": [],
                })
        else:
            review.append({
                "key": key,
                "reason": "This initial + last name is not unique, so HR must verify the correct employee/file.",
                "employees": employee_group,
                "files": [item["filename"] for item in file_group],
            })
            used_file_keys.add(key)

    unmatched_files = []
    for key, file_group in files_by_key.items():
        if key not in employees_by_key:
            unmatched_files.extend(item["filename"] for item in file_group)

    found.sort(key=lambda item: (item["employee"]["last"].casefold(), item["employee"]["first"].casefold()))
    missing.sort(key=lambda employee: (employee["last"].casefold(), employee["first"].casefold()))
    review.sort(key=lambda item: item["key"])
    unmatched_files.sort(key=str.casefold)

    return found, missing, review, unmatched_files
