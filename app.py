from flask import Flask, render_template, request, send_file
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor

import csv
import os
import tempfile
import zipfile
from datetime import datetime

import pdfplumber
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

from tools.lunch_review import extract_employee_page, hours_between
from tools.attendance_review import build_daily_review
from tools.attendance_pdf import create_attendance_pdf
from tools.ad347_data import build_ad347_record
from tools.pdf_generator import create_ad347_pdf
from tools.talent_lms import (
    build_incomplete_list,
    find_possible_name_mismatches,
    read_talent_lms,
    read_ukg_employees,
)


app = Flask(__name__)

TEMPLATE_PATH = "BLANK_AD347.pdf"


@app.route("/attendance-review", methods=["GET", "POST"])
def attendance_review():
    if request.method == "GET":
        return render_template("attendance_review.html")

    exception_file = request.files.get("exception_file")
    lunch_file = request.files.get("lunch_file")

    if not exception_file or not lunch_file:
        return "<h2>Error</h2><p>Both Excel reports are required.</p>", 400

    temp_folder = tempfile.mkdtemp()
    exception_path = os.path.join(temp_folder, "exception_report.xlsx")
    lunch_path = os.path.join(temp_folder, "lunch_report.xlsx")
    exception_file.save(exception_path)
    lunch_file.save(lunch_path)

    try:
        report_date, rows, missing_punches = build_daily_review(
            exception_path,
            lunch_path,
        )

        if missing_punches:
            names = "".join(f"<li>{name}</li>" for name in missing_punches)
            return f"""
                <h2>Fix missing punches first</h2>
                <p>The report was stopped because these employees still have a missing punch:</p>
                <ul>{names}</ul>
                <p>Correct them in UKG, rerun both reports, and upload them again.</p>
                <p><a href='/attendance-review'>Try again</a></p>
            """, 400

        output_path = os.path.join(
            temp_folder,
            "Daily_Attendance_Exception_Report.pdf",
        )
        create_attendance_pdf(report_date, rows, output_path)

        return send_file(
            output_path,
            as_attachment=True,
            download_name=(
                f"Daily_Attendance_Exceptions_{report_date.isoformat()}.pdf"
            ),
            mimetype="application/pdf",
        )
    except Exception as error:
        return f"""
            <h2>Error</h2>
            <p>{error}</p>
            <p><a href='/attendance-review'>Try again</a></p>
        """, 400


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


def _talent_department(employee):
    department = employee["department"].strip().upper()
    job = employee["job"].strip().upper()
    if department == "NSG" and job == "CNA":
        return "CNAs"
    if department == "NSG":
        return "Nursing"
    return employee["department"] or "No Department"


def create_talent_lms_docx(incomplete, month_label, eligible_count, completed_count, output_path):
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.right_margin = Inches(0.85)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)

    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.1

    title = document.add_paragraph()
    title.paragraph_format.space_after = Pt(4)
    title_run = title.add_run(f"{month_label} In-Service - Not Completed")
    title_run.bold = True
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(18)
    title_run.font.color.rgb = RGBColor(31, 78, 120)

    summary = document.add_paragraph()
    summary.paragraph_format.space_after = Pt(12)
    summary_run = summary.add_run(
        f"Eligible: {eligible_count}    Completed: {completed_count}    Not completed: {len(incomplete)}"
    )
    summary_run.bold = True

    note = document.add_paragraph(
        "Editable follow-up list. Delete any employee who should not be included before printing."
    )
    note.paragraph_format.space_after = Pt(14)

    if not incomplete:
        document.add_paragraph("Everyone required for this month has a completed record.")
    else:
        sorted_employees = sorted(
            incomplete,
            key=lambda employee: (
                _talent_department(employee),
                employee["last"],
                employee["first"],
            ),
        )
        current_department = None
        for employee in sorted_employees:
            department = _talent_department(employee)
            if department != current_department:
                heading = document.add_paragraph()
                heading.paragraph_format.keep_with_next = True
                heading.paragraph_format.space_before = Pt(10)
                heading.paragraph_format.space_after = Pt(4)
                heading_run = heading.add_run(department)
                heading_run.bold = True
                heading_run.font.name = "Calibri"
                heading_run.font.size = Pt(13)
                heading_run.font.color.rgb = RGBColor(31, 78, 120)
                current_department = department

            paragraph = document.add_paragraph()
            paragraph.paragraph_format.left_indent = Inches(0.15)
            paragraph.paragraph_format.space_after = Pt(3)
            paragraph.add_run(employee["display_name"])

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.add_run("For supervisor follow-up. UKG employment dates determine monthly eligibility.")
    footer_run.font.name = "Calibri"
    footer_run.font.size = Pt(8)

    document.save(output_path)


@app.route("/talent-lms", methods=["GET", "POST"])
def talent_lms():
    if request.method == "GET":
        return render_template("talent_lms.html")

    uploaded_files = request.files.getlist("report_files")
    month_value = request.form.get("month", "")

    ukg_file = next(
        (file for file in uploaded_files if file.filename.lower().endswith(".xlsx")),
        None,
    )
    talent_file = next(
        (file for file in uploaded_files if file.filename.lower().endswith(".csv")),
        None,
    )
    if not ukg_file or not talent_file:
        return "<h2>Error</h2><p>Drop one UKG Excel file and one TalentLMS CSV file into the box.</p><p><a href='/talent-lms'>Try again</a></p>", 400

    try:
        year_text, month_text = month_value.split("-", 1)
        year = int(year_text)
        month = int(month_text)
        month_date = datetime(year, month, 1)
    except (ValueError, AttributeError):
        return "<h2>Error</h2><p>Select the in-service month.</p><p><a href='/talent-lms'>Try again</a></p>", 400

    temp_folder = tempfile.mkdtemp()
    ukg_path = os.path.join(temp_folder, "employees.xlsx")
    talent_path = os.path.join(temp_folder, "talent.csv")
    ukg_file.save(ukg_path)
    talent_file.save(talent_path)

    try:
        employees = read_ukg_employees(ukg_path, year, month)
        talent_records = read_talent_lms(talent_path)
        mismatches = find_possible_name_mismatches(employees, talent_records)
        incomplete, completed_count = build_incomplete_list(employees, talent_records)
        month_label = month_date.strftime("%B %Y")

        if mismatches:
            rows = "".join(
                f"<tr><td>{item['ukg_name']}</td><td>{item['talent_name']}</td><td>{item['talent_status']}</td></tr>"
                for item in mismatches
            )
            return f"""
            <h2>{month_label} — Name Review</h2>
            <div style="background:#fff3cd;padding:14px;max-width:760px">
                <b>Possible name mismatch found.</b> Correct it in TalentLMS if these are the same person,
                export the course report again, and create the list again.
            </div>
            <table border="1" cellpadding="8" cellspacing="0" style="margin-top:18px;border-collapse:collapse">
                <tr><th>UKG Name</th><th>TalentLMS Name</th><th>TalentLMS Status</th></tr>
                {rows}
            </table>
            <p><a href="/talent-lms">Try again</a> | <a href="/">Home</a></p>
            """

        docx_path = os.path.join(temp_folder, "TalentLMS_Not_Completed.docx")
        create_talent_lms_docx(
            incomplete,
            month_label,
            len(employees),
            completed_count,
            docx_path,
        )
        safe_month = month_date.strftime("%Y-%m")
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=f"TalentLMS_{safe_month}_Not_Completed.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as error:
        return f"<h2>Error</h2><pre>{error}</pre><p><a href='/talent-lms'>Try again</a></p>", 400


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
