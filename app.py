from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

import csv
import os
import tempfile
import zipfile

import pdfplumber

from tools.lunch_review import extract_employee_page, hours_between
from tools.ad347_data import build_ad347_record
from tools.pdf_generator import create_ad347_pdf


app = Flask(__name__)

TEMPLATE_PATH = "BLANK_AD347.pdf"


@app.route("/")
def home():
    return render_template("index.html")


def review_timesheets(pdf_path):
    grouped_records = {}
    preview_rows = []

    with pdfplumber.open(pdf_path) as document:
        for page in document.pages:
            text = page.extract_text() or ""

            name, department, shifts = extract_employee_page(text)

            missed_dates = []

            for work_date, segments in shifts.items():
                total_hours = sum(
                    hours_between(start, end)
                    for start, end in segments
                )

                if total_hours < 7:
                    continue

                lunch_taken = len(segments) >= 2

                preview_rows.append(
                    {
                        "name": name,
                        "department": department,
                        "date": work_date,
                        "hours": total_hours,
                        "lunch_taken": lunch_taken,
                    }
                )

                if not lunch_taken:
                    missed_dates.append(work_date)

            if missed_dates:
                key = (name, department)
                grouped_records.setdefault(key, [])
                grouped_records[key].extend(missed_dates)

    return grouped_records, preview_rows


def build_preview_page(preview_rows):
    html = """
    <h2>Lunch Review Preview</h2>
    <p><b>No forms were created.</b></p>
    """

    if not preview_rows:
        html += "<p>No qualifying shifts were found.</p>"
    else:
        current_employee = None

        for row in preview_rows:
            employee_key = (row["name"], row["department"])

            if employee_key != current_employee:
                html += f"""
                <hr>
                <h3>{row["name"]}</h3>
                <p><b>Department:</b> {row["department"]}</p>
                """
                current_employee = employee_key

            status = (
                "✅ Lunch Taken"
                if row["lunch_taken"]
                else "❌ No Recorded Lunch"
            )

            html += (
                f'{row["date"]} — '
                f'{row["hours"]:.2f} hours — '
                f'{status}<br>'
            )

    html += """
    <br>
    <p><a href="/lunch-review">Run another review</a></p>
    <p><a href="/">Home</a></p>
    """

    return html


def create_supervisor_pdf(grouped_records, output_path):
    pdf = canvas.Canvas(output_path, pagesize=letter)

    page_width, page_height = letter
    y = page_height - 55

    pdf.setTitle("Supervisor Missed Lunch List")
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(50, y, "Supervisor Missed Lunch List")

    y -= 22
    pdf.setFont("Helvetica", 10)
    pdf.drawString(
        50,
        y,
        "Employees with no recorded meal break on qualifying shifts"
    )

    y -= 30

    sorted_records = sorted(
        grouped_records.items(),
        key=lambda item: (item[0][1], item[0][0]),
    )

    current_department = None

    for (name, department), dates in sorted_records:
        unique_dates = sorted(set(dates))

        if department != current_department:
            if y < 120:
                pdf.showPage()
                y = page_height - 55

            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(50, y, department)
            y -= 20
            current_department = department

        if y < 100:
            pdf.showPage()
            y = page_height - 55

            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(50, y, department)
            y -= 20

        pdf.setFont("Helvetica-Bold", 10)
        pdf.drawString(65, y, name)

        y -= 15
        pdf.setFont("Helvetica", 10)
        pdf.drawString(
            80,
            y,
            "Missed lunch date(s): " + ", ".join(unique_dates)
        )

        y -= 15
        pdf.drawString(
            80,
            y,
            f"Total missed lunches: {len(unique_dates)}"
        )

        y -= 22

    pdf.setFont("Helvetica-Oblique", 8)
    pdf.drawString(
        50,
        35,
        "This list is for supervisor follow-up during the pay period."
    )

    pdf.save()


def create_final_zip(grouped_records, temp_folder):
    pdf_files = []
    summary_rows = []

    sorted_records = sorted(
        grouped_records.items(),
        key=lambda item: item[0][0],
    )

    for index, ((name, department), dates) in enumerate(
        sorted_records,
        start=1,
    ):
        unique_dates = sorted(set(dates))

        record = build_ad347_record(
            name=name,
            department=department,
            dates=unique_dates,
        )

        safe_name = (
            name.replace(" ", "_")
            .replace("/", "-")
            .replace("\\", "-")
        )

        output_path = os.path.join(
            temp_folder,
            f"{index}_{safe_name}_AD347.pdf",
        )

        create_ad347_pdf(
            record=record,
            template_path=TEMPLATE_PATH,
            output_path=output_path,
        )

        pdf_files.append(output_path)

        summary_rows.append(
            {
                "Employee Name": name,
                "Department": department,
                "Missed Lunch Dates": ", ".join(unique_dates),
                "Total Missed Lunches": len(unique_dates),
            }
        )

    summary_path = os.path.join(
        temp_folder,
        "Lunch_Review_Summary.csv",
    )

    with open(
        summary_path,
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as summary_file:
        fieldnames = [
            "Employee Name",
            "Department",
            "Missed Lunch Dates",
            "Total Missed Lunches",
        ]

        writer = csv.DictWriter(
            summary_file,
            fieldnames=fieldnames,
        )

        writer.writeheader()
        writer.writerows(summary_rows)

    zip_path = os.path.join(
        temp_folder,
        "AD347_Forms_and_Summary.zip",
    )

    with zipfile.ZipFile(zip_path, "w") as zip_file:
        zip_file.write(
            summary_path,
            arcname="Lunch_Review_Summary.csv",
        )

        for pdf_file in pdf_files:
            zip_file.write(
                pdf_file,
                arcname=os.path.basename(pdf_file),
            )

    return zip_path


@app.route("/lunch-review", methods=["GET", "POST"])
def lunch_review():
    if request.method == "POST":
        pdf = request.files.get("pdf")
        action = request.form.get("action")

        if not pdf or pdf.filename == "":
            return "No file selected"

        temp_folder = tempfile.mkdtemp()
        temp_pdf = os.path.join(temp_folder, "timesheets.pdf")
        pdf.save(temp_pdf)

        try:
            grouped_records, preview_rows = review_timesheets(temp_pdf)

            if action == "preview":
                return build_preview_page(preview_rows)

            if action == "supervisor":
                if not grouped_records:
                    return """
                    <h2>No missed lunches found.</h2>
                    <p><a href="/lunch-review">Run another review</a></p>
                    """

                supervisor_path = os.path.join(
                    temp_folder,
                    "Supervisor_Missed_Lunch_List.pdf",
                )

                create_supervisor_pdf(
                    grouped_records,
                    supervisor_path,
                )

                return send_file(
                    supervisor_path,
                    as_attachment=True,
                    download_name="Supervisor_Missed_Lunch_List.pdf",
                )

            if action == "final":
                if not grouped_records:
                    return """
                    <h2>No AD-347 forms needed.</h2>
                    <p><a href="/lunch-review">Run another review</a></p>
                    """

                zip_path = create_final_zip(
                    grouped_records,
                    temp_folder,
                )

                return send_file(
                    zip_path,
                    as_attachment=True,
                    download_name="AD347_Forms_and_Summary.zip",
                )

            return "No action selected"

        except Exception as error:
            return f"<h2>Error</h2><pre>{error}</pre>"

    return """
    <h2>Lunch Review</h2>

    <form method="POST" enctype="multipart/form-data">
        <input
            type="file"
            name="pdf"
            accept=".pdf"
            required
        >

        <br><br>

        <button type="submit" name="action" value="preview">
            Preview Lunch Review
        </button>

        <br><br>

        <button type="submit" name="action" value="supervisor">
            Print Supervisor List
        </button>

        <br><br>

        <button type="submit" name="action" value="final">
            Generate Final AD-347 Forms
        </button>
    </form>

    <p>
        Use Preview or Supervisor List during the pay period.
        Use Final only after the pay period is complete.
    </p>

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
