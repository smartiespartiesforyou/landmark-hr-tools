from flask import Flask, render_template

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/lunch-review")
def lunch_review():
    return "<h2>Lunch Review - Coming Soon</h2>"

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
