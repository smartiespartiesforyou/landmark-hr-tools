from flask import Flask, render_template, request
from datetime import datetime, timedelta
import os
import re
import tempfile
import pdfplumber

app = Flask(__name__)


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
        re.DOTALL
    )

    department_match = re.search(r"Jobs \(HR\)\s+(.+)", text)

    name = name_match.group(1).strip() if name_match else "Unknown Employee"
    department = department_match.group(1).strip() if department_match else ""

    shifts = {}
    current_date = None

    for line in text.splitlines():
        line = line.strip()

        if line.startswith("Week Total:") or line.startswith("Total:"):
            break

        dated_row = re.match(
            r"^[A-Z][a-z]{2}\s+"
            r"(\d{2}/\d{2}/\d{4})\s+"
            r"(\d{2}:\d{2}[ap])\s+"
            r"(?:[A-Z][a-z]{2}\s+)?"
            r"(\d{2}:\d{2}[ap])\b",
            line
        )

        continuation_row = re.match(
            r"^(\d{2}:\d{2}[ap])\s+"
            r"(?:[A-Z][a-z]{2}\s+)?"
            r"(\d{2}:\d{2}[ap])\b",
            line
        )

        if dated_row:
            current_date = dated_row.group(1)

            shifts.setdefault(current_date, []).append(
                (dated_row.group(2), dated_row.group(3))
            )

        elif continuation_row and current_date:
            shifts.setdefault(current_date, []).append(
                (continuation_row.group(1), continuation_row.group(2))
            )

    return name, department, shifts


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/lunch-review", methods=["GET", "POST"])
def lunch_review():
    if request.method == "POST":
        pdf = request.files.get("pdf")

        if not pdf or pdf.filename == "":
            return "No file selected"

        temp_path = None

        try:
            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".pdf"
            ) as temp_file:
                pdf.save(temp_file.name)
                temp_path = temp_file.name

            html = "<h2>Lunch Review</h2>"
            no_lunch_count = 0

            with pdfplumber.open(temp_path) as document:
                for page in document.pages:
                    text = page.extract_text() or ""

                    name, department, shifts = extract_employee_page(text)

                    department_upper = department.upper()

                    if "LPN" in department_upper or "RGN" in department_upper:
                        continue

                    employee_results = []

                    for work_date, segments in shifts.items():
                        total_hours = sum(
                            hours_between(start, end)
                            for start, end in segments
                        )

                        if total_hours < 7:
                            continue

                        if len(segments) >= 2:
                            result = "✅ Lunch Taken"
                        else:
                            result = "❌ No Lunch"
                            no_lunch_count += 1

                        employee_results.append(
                            (work_date, total_hours, result)
                        )

                    if employee_results:
                        html += f"<hr><h3>{name}</h3>"
                        html += f"<b>Department:</b> {department}<br><br>"

                        for work_date, total_hours, result in employee_results:
                            html += (
                                f"{work_date} — "
                                f"{total_hours:.2f} hours — "
                                f"{result}<br>"
                            )

            html += f"<hr><h3>AD-347 dates found: {no_lunch_count}</h3>"
            html += '<p><a href="/lunch-review">Upload another PDF</a></p>'

            return html

        except Exception as error:
            return f"<h2>Error</h2><pre>{error}</pre>"

        finally:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)

    return """
    <h2>Lunch Review</h2>

    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="pdf" accept=".pdf" required>
        <br><br>
        <input type="submit" value="Run Lunch Review">
    </form>

    <p><a href="/">Home</a></p>
    """


@app.route("/ad347")
def ad347():
    return "<h2>Coming Soon</h2>"


@app.route("/evaluations")
def evaluations():
    return "<h2>Coming Soon</h2>"


@app.route("/anniversaries")
def anniversaries():
    return "<h2>Coming Soon</h2>"


if __name__ == "__main__":
    app.run(debug=True)
