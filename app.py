from flask import Flask, render_template, request, send_file
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


@app.route("/lunch-review", methods=["GET", "POST"])
def lunch_review():
    if request.method == "POST":
        pdf = request.files.get("pdf")

        if not pdf or pdf.filename == "":
            return "No file selected"

        temp_folder = tempfile.mkdtemp()
        temp_pdf = os.path.join(temp_folder, "timesheets.pdf")
        pdf.save(temp_pdf)

        try:
            grouped_records = {}

            with pdfplumber.open(temp_pdf) as document:
                for page in document.pages:
                    text = page.extract_text() or ""

                    name, department, shifts = extract_employee_page(text)

                    department_upper = department.upper()

                    if "LPN" in department_upper or "RGN" in department_upper:
                        continue

                    missed_dates = []

                    for work_date, segments in shifts.items():
                        total_hours = sum(
                            hours_between(start, end)
                            for start, end in segments
                        )

                        if total_hours < 7:
                            continue

                        if len(segments) == 1:
                            missed_dates.append(work_date)

                    if missed_dates:
                        key = (name, department)

                        if key not in grouped_records:
                            grouped_records[key] = []

                        grouped_records[key].extend(missed_dates)

            if not grouped_records:
                return """
                <h2>No AD-347 forms needed.</h2>
                <p><a href="/lunch-review">Run another review</a></p>
                """

            pdf_files = []

            for index, ((name, department), dates) in enumerate(
                grouped_records.items(),
                start=1,
            ):
                unique_dates = sorted(set(dates))

                record = build_ad347_record(
                    name=name,
                    department=department,
                    dates=unique_dates,
                )

                safe_name = name.replace(" ", "_")

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

            zip_path = os.path.join(temp_folder, "AD347_Forms.zip")

            with zipfile.ZipFile(zip_path, "w") as zip_file:
                for pdf_file in pdf_files:
                    zip_file.write(
                        pdf_file,
                        arcname=os.path.basename(pdf_file),
                    )

            return send_file(
                zip_path,
                as_attachment=True,
                download_name="AD347_Forms.zip",
            )

        except Exception as error:
            return f"<h2>Error</h2><pre>{error}</pre>"

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
