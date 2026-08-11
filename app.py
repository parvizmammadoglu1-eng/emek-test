
from flask import Flask, render_template, request, redirect, url_for, session
import sqlite3
from datetime import datetime
from pathlib import Path

app = Flask(__name__)
app.secret_key = "CHANGE_THIS_SECRET_KEY"
DB = Path(__file__).with_name("test.db")

QUESTIONS = [
    ("Əmək müqaviləsi hansı formada bağlanır?", "A) Şifahi", "B) Yazılı", "C) İstənilən formada", "D) Heç biri", "B"),
    ("Əmək münasibətlərini əsasən hansı sənəd tənzimləyir?", "A) Mülki Məcəllə", "B) Vergi Məcəlləsi", "C) Əmək Məcəlləsi", "D) Konstitusiya", "C"),
    ("İşçinin əsas hüquqlarından biri hansıdır?", "A) Əmək haqqı almaq", "B) İşə gəlməmək", "C) Qaydaları pozmaq", "D) Müqaviləsiz işləmək", "A"),
]

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def init_db():
    con = db()
    con.execute("""CREATE TABLE IF NOT EXISTS questions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        question TEXT NOT NULL, a TEXT NOT NULL, b TEXT NOT NULL,
        c TEXT NOT NULL, d TEXT NOT NULL, answer TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS results (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, correct INTEGER NOT NULL, total INTEGER NOT NULL,
        percent INTEGER NOT NULL, created_at TEXT NOT NULL)""")
    if con.execute("SELECT COUNT(*) FROM questions").fetchone()[0] == 0:
        con.executemany(
            "INSERT INTO questions(question,a,b,c,d,answer) VALUES (?,?,?,?,?,?)",
            QUESTIONS
        )
    con.commit()
    con.close()

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        if not name:
            return render_template("home.html", error="Ad və soyad daxil edin.")
        session["name"] = name
        session["answers"] = {}
        return redirect(url_for("question", n=0))
    return render_template("home.html")

@app.route("/question/<int:n>", methods=["GET", "POST"])
def question(n):
    con = db()
    questions = con.execute("SELECT * FROM questions ORDER BY id").fetchall()
    con.close()
    if "name" not in session:
        return redirect(url_for("home"))
    if n >= len(questions):
        return redirect(url_for("finish"))
    if request.method == "POST":
        selected = request.form.get("answer")
        if selected not in ["A","B","C","D"]:
            return redirect(url_for("question", n=n))
        answers = session.get("answers", {})
        answers[str(n)] = selected
        session["answers"] = answers
        return redirect(url_for("question", n=n+1))
    return render_template("question.html", q=questions[n], n=n, total=len(questions), name=session["name"])

@app.route("/finish")
def finish():
    if "name" not in session:
        return redirect(url_for("home"))
    con = db()
    questions = con.execute("SELECT * FROM questions ORDER BY id").fetchall()
    con.close()
    answers = session.get("answers", {})
    correct = sum(1 for i,q in enumerate(questions) if answers.get(str(i)) == q["answer"])
    total = len(questions)
    percent = round(correct / total * 100) if total else 0
    con = db()
    con.execute(
        "INSERT INTO results(name,correct,total,percent,created_at) VALUES (?,?,?,?,?)",
        (session["name"], correct, total, percent, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    )
    con.commit(); con.close()
    name = session["name"]
    session.clear()
    return render_template("finish.html", name=name, correct=correct, total=total, percent=percent)

@app.route("/admin")
def admin():
    con = db()
    results = con.execute("SELECT * FROM results ORDER BY id DESC").fetchall()
    con.close()
    return render_template("admin.html", results=results)

if __name__ == "__main__":
    init_db()
    app.run(debug=True)
