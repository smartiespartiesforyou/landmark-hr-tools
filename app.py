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

        pdf = request.files["pdf"]

        filepath = os.path.join(UPLOAD_FOLDER, pdf.filename)
        pdf.save(filepath)

        html = "<h2>Lunch Review</h2>"

        with pdfplumber.open(filepath) as document:

            for page in document.pages:

                text = page.extract_text() or ""

                name = re.search(
                    r"Employee Timesheet.*?\n(.+?)\s+\(Employee Id:",
                    text,
                    re.DOTALL
                )

                if name:
                    html += f"<hr><h3>{name.group(1).strip()}</h3>"

                dates = re.findall(r"\d{2}/\d{2}/\d{4}", text)

                counts = {}

                for d in dates:
                    counts[d] = counts.get(d, 0) + 1

                for d, c in counts.items():

                    if c == 1:
                        result = "❌ No Lunch"

                    elif c == 2:
                        result = "✅ Lunch Taken"

                    else:
                        result = "🟡 Review"

                    html += f"{d} — {result}<br>"

        os.remove(filepath)

        return html

    return """
    <h2>Lunch Review</h2>

    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="pdf" accept=".pdf"><br><br>
        <input type="submit" value="Upload PDF">
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
