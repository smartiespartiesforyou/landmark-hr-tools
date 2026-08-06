from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return "<h1>Landmark HR Tools</h1><p>Server Running ✅</p>"

if __name__ == "__main__":
    app.run(debug=True)
