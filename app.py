import os
import sqlite3
from datetime import datetime
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, session

app = Flask(__name__)

# Render-də Environment Variable-dan götürülür
app.secret_key = os.environ.get(
    "SECRET_KEY",
    "temporary-secret-key-change-me"
)

DB = Path(__file__).with_name("test.db")

QUESTIONS = [
    (
        "Əmək müqaviləsi hansı formada bağlanır?",
        "A) Şifahi",
        "B) Yazılı",
        "C) İstənilən formada",
        "D) Heç biri",
        "B"
    ),
    (
        "Əmək münasibətlərini əsasən hansı sənəd tənzimləyir?",
        "A) Mülki Məcəllə",
        "B) Vergi Məcəlləsi",
        "C) Əmək Məcəlləsi",
        "D) Konstitusiya",
        "C"
    ),
    (
        "İşçinin əsas hüquqlarından biri hansıdır?",
        "A) Əmək haqqı almaq",
        "B) İşə gəlməmək",
        "C) Qaydaları pozmaq",
        "D) Müqaviləsiz işləmək",
        "A"
    ),
]


def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()

    con.execute("""
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question TEXT NOT NULL,
            a TEXT NOT NULL,
            b TEXT NOT NULL,
            c TEXT NOT NULL,
            d TEXT NOT NULL,
            answer TEXT NOT NULL
        )
    """)

    con.execute("""
        CREATE TABLE IF NOT EXISTS results (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            correct INTEGER NOT NULL,
            total INTEGER NOT NULL,
            percent INTEGER NOT NULL,
            created_at TEXT NOT NULL
        )
    """)

    count = con.execute(
        "SELECT COUNT(*) FROM questions"
    ).fetchone()[0]

    if count == 0:
        con.executemany(
            """
            INSERT INTO questions
            (question, a, b, c, d, answer)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            QUESTIONS
        )

    con.commit()
    con.close()


# Tətbiq Render/Gunicorn ilə açılanda da bazanı yaradır
init_db()


@app.route("/", methods=["GET", "POST"])
def home():

    if request.method == "POST":

        name = request.form.get("name", "").strip()

        if not name:
            return render_template(
                "home.html",
                error="Ad və soyad daxil edin."
            )

        session["name"] = name
        session["answers"] = {}

        return redirect(url_for("question", n=0))

    return render_template("home.html")


@app.route("/question/<int:n>", methods=["GET", "POST"])
def question(n):

    if "name" not in session:
        return redirect(url_for("home"))

    con = db()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
    ).fetchall()

    con.close()

    if n >= len(questions):
        return redirect(url_for("finish"))

    if request.method == "POST":

        selected = request.form.get("answer")

        if selected not in ["A", "B", "C", "D"]:
            return redirect(
                url_for("question", n=n)
            )

        answers = session.get("answers", {})

        answers[str(n)] = selected

        session["answers"] = answers

        return redirect(
            url_for("question", n=n + 1)
        )

    return render_template(
        "question.html",
        q=questions[n],
        n=n,
        total=len(questions),
        name=session["name"]
    )


@app.route("/finish")
def finish():

    if "name" not in session:
        return redirect(url_for("home"))

    con = db()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
    ).fetchall()

    con.close()

    answers = session.get("answers", {})

    correct = sum(
        1
        for i, q in enumerate(questions)
        if answers.get(str(i)) == q["answer"]
    )

    total = len(questions)

    percent = round(
        correct / total * 100
    ) if total else 0

    name = session["name"]

    con = db()

    con.execute(
        """
        INSERT INTO results
        (name, correct, total, percent, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            name,
            correct,
            total,
            percent,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )
    )

    con.commit()
    con.close()

    session.clear()

    return render_template(
        "finish.html",
        name=name,
        correct=correct,
        total=total,
        percent=percent
    )


@app.route("/admin")
def admin():

    con = db()

    results = con.execute(
        "SELECT * FROM results ORDER BY id DESC"
    ).fetchall()

    con.close()

    return render_template(
        "admin.html",
        results=results
    )
