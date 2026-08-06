from flask import Flask, render_template, request
import os
import re
import pdfplumber

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/lunch-review", methods=["GET", "POST"])
def lunch_review():
    if request.method == "POST":
        if "pdf" not in request.files:
            return "No file selected"

        pdf = request.files["pdf"]

        if pdf.filename == "":
            return "No file selected"

        filepath = os.path.join(UPLOAD_FOLDER, pdf.filename)
        pdf.save(filepath)

        with pdfplumber.open(filepath) as document:
            first_page_text = document.pages[0].extract_text() or ""

        match = re.search(
            r"Employee Timesheet.*?\n(.+?)\s+\(Employee Id:",
            first_page_text,
            re.DOTALL
        )

        employee_name = match.group(1).strip() if match else "Employee not found"

        os.remove(filepath)

        return f"""
        <h2>PDF Read Successfully</h2>
        <p>First employee: <strong>{employee_name}</strong></p>
        <p><a href="/lunch-review">Upload another PDF</a></p>
        """

    return """
    <!DOCTYPE html>
    <html>
    <body>

    <h2>Lunch Review</h2>

    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="pdf" accept=".pdf">
        <br><br>
        <input type="submit" value="Upload PDF">
    </form>

    <p><a href="/">Back to Home</a></p>

    </body>
    </html>
    """


@app.route("/ad347")
def ad347():
    return "<h2>AD-347 Generator - Coming Soon</h2>"


@app.route("/evaluations")
def evaluations():
    return "<h2>Evaluation Finder - Coming Soon</h2>"


@app.route("/anniversaries")
def anniversaries():
    return "<h2>Anniversary Finder - Coming Soon</h2>"


if __name__ == "__main__":
    app.run(debug=True)
