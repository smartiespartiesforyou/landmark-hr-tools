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

        pdf = request.files["pdf"]

        if pdf.filename != "":
            pdf.save(os.path.join(UPLOAD_FOLDER, pdf.filename))
            return "<h2>Upload Successful!</h2>"

    return """
    <h2>Lunch Review</h2>

    <form method="POST" enctype="multipart/form-data">
        <input type="file" name="pdf">
        <br><br>
        <button type="submit">Upload PDF</button>
    </form>
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
