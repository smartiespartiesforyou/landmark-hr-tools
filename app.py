from flask import Flask, render_template, request
import os

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

        return f"<h2>Upload Successful!</h2><p>{pdf.filename}</p>"

    return """
<!DOCTYPE html>
<html>
<body>

<h2>Lunch Review</h2>

<form method="POST" enctype="multipart/form-data">
    <input type="file" name="pdf">
    <br><br>
    <input type="submit" value="Upload PDF">
</form>

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
