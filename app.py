import os
from datetime import datetime
from io import BytesIO
from functools import wraps

import psycopg2
from psycopg2.extras import RealDictCursor

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    send_file
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


app = Flask(__name__)

# ==========================================================
# SETTINGS
# ==========================================================

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "temporary-secret-key-change-me"
)

ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "12345"
)

DATABASE_URL = os.environ.get("DATABASE_URL")


# ==========================================================
# QUESTIONS
# ==========================================================

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


# ==========================================================
# DATABASE CONNECTION
# ==========================================================

def db():
    if not DATABASE_URL:
        raise RuntimeError(
            "DATABASE_URL Render Environment Variables bölməsində "
            "əlavə edilməyib."
        )

    return psycopg2.connect(
        DATABASE_URL,
        cursor_factory=RealDictCursor
    )


# ==========================================================
# DATABASE INITIALIZATION
# ==========================================================

def init_db():

    con = db()

    cur = con.cursor()

    # QUESTIONS TABLE
    cur.execute("""
        CREATE TABLE IF NOT EXISTS questions (
            id SERIAL PRIMARY KEY,
            question TEXT NOT NULL,
            a TEXT NOT NULL,
            b TEXT NOT NULL,
            c TEXT NOT NULL,
            d TEXT NOT NULL,
            answer TEXT NOT NULL
        )
    """)

    # RESULTS TABLE
    cur.execute("""
        CREATE TABLE IF NOT EXISTS results (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            correct INTEGER NOT NULL,
            total INTEGER NOT NULL,
            percent INTEGER NOT NULL,
            created_at TIMESTAMP NOT NULL
        )
    """)

    # Əgər suallar bazada yoxdursa,
    # ilkin sualları əlavə et
    cur.execute(
        "SELECT COUNT(*) AS count FROM questions"
    )

    count = cur.fetchone()["count"]

    if count == 0:

        for question in QUESTIONS:

            cur.execute(
                """
                INSERT INTO questions
                (question, a, b, c, d, answer)
                VALUES (%s, %s, %s, %s, %s, %s)
                """,
                question
            )

    con.commit()

    cur.close()
    con.close()


# ==========================================================
# INITIALIZE DATABASE
# ==========================================================

if DATABASE_URL:
    init_db()


# ==========================================================
# ADMIN AUTHENTICATION
# ==========================================================

def admin_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("admin"):
            return redirect(
                url_for("admin_login")
            )

        return function(*args, **kwargs)

    return wrapper


# ==========================================================
# HOME
# ==========================================================

@app.route("/", methods=["GET", "POST"])
def home():

    if request.method == "POST":

        name = request.form.get(
            "name",
            ""
        ).strip()

        if not name:

            return render_template(
                "home.html",
                error="Ad və soyad daxil edin."
            )

        session["name"] = name

        session["answers"] = {}

        return redirect(
            url_for(
                "question",
                n=0
            )
        )

    return render_template(
        "home.html"
    )


# ==========================================================
# QUESTIONS
# ==========================================================

@app.route(
    "/question/<int:n>",
    methods=["GET", "POST"]
)
def question(n):

    if "name" not in session:

        return redirect(
            url_for("home")
        )

    con = db()

    cur = con.cursor()

    cur.execute(
        """
        SELECT *
        FROM questions
        ORDER BY id
        """
    )

    questions = cur.fetchall()

    cur.close()
    con.close()

    if n >= len(questions):

        return redirect(
            url_for("finish")
        )

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        if selected not in [
            "A",
            "B",
            "C",
            "D"
        ]:

            return redirect(
                url_for(
                    "question",
                    n=n
                )
            )

        answers = session.get(
            "answers",
            {}
        )

        answers[str(n)] = selected

        session["answers"] = answers

        return redirect(
            url_for(
                "question",
                n=n + 1
            )
        )

    return render_template(
        "question.html",
        q=questions[n],
        n=n,
        total=len(questions),
        name=session["name"]
    )


# ==========================================================
# FINISH
# ==========================================================

@app.route("/finish")
def finish():

    if "name" not in session:

        return redirect(
            url_for("home")
        )

    con = db()

    cur = con.cursor()

    cur.execute(
        """
        SELECT *
        FROM questions
        ORDER BY id
        """
    )

    questions = cur.fetchall()

    answers = session.get(
        "answers",
        {}
    )

    correct = sum(
        1
        for i, q in enumerate(questions)
        if answers.get(str(i)) == q["answer"]
    )

    total = len(questions)

    percent = (
        round(correct / total * 100)
        if total
        else 0
    )

    name = session["name"]

    # NƏTİCƏNİ POSTGRESQL-Ə YAZ
    cur.execute(
        """
        INSERT INTO results
        (
            name,
            correct,
            total,
            percent,
            created_at
        )
        VALUES (%s, %s, %s, %s, %s)
        """,
        (
            name,
            correct,
            total,
            percent,
            datetime.now()
        )
    )

    con.commit()

    cur.close()
    con.close()

    session.clear()

    return render_template(
        "finish.html",
        name=name,
        correct=correct,
        total=total,
        percent=percent
    )


# ==========================================================
# ADMIN LOGIN
# ==========================================================

@app.route(
    "/admin/login",
    methods=["GET", "POST"]
)
def admin_login():

    if request.method == "POST":

        password = request.form.get(
            "password",
            ""
        )

        if password == ADMIN_PASSWORD:

            session["admin"] = True

            return redirect(
                url_for("admin")
            )

        return render_template(
            "admin_login.html",
            error="Parol yanlışdır."
        )

    return render_template(
        "admin_login.html"
    )


# ==========================================================
# ADMIN LOGOUT
# ==========================================================

@app.route("/admin/logout")
def admin_logout():

    session.pop(
        "admin",
        None
    )

    return redirect(
        url_for("admin_login")
    )


# ==========================================================
# ADMIN PANEL
# ==========================================================

@app.route("/admin")
@admin_required
def admin():

    con = db()

    cur = con.cursor()

    # NƏTİCƏLƏR
    cur.execute(
        """
        SELECT *
        FROM results
        ORDER BY id DESC
        """
    )

    results = cur.fetchall()

    # SUALLAR
    cur.execute(
        """
        SELECT *
        FROM questions
        ORDER BY id
        """
    )

    questions = cur.fetchall()

    cur.close()
    con.close()

    return render_template(
        "admin.html",
        results=results,
        questions=questions
    )


# ==========================================================
# EXCEL EXPORT
# ==========================================================

@app.route("/admin/export")
@admin_required
def admin_export():

    con = db()

    cur = con.cursor()

    cur.execute(
        """
        SELECT *
        FROM results
        ORDER BY id DESC
        """
    )

    results = cur.fetchall()

    cur.close()
    con.close()

    # EXCEL YARAT
    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Test nəticələri"

    # BAŞLIQLAR
    headers = [
        "№",
        "Ad və soyad",
        "Düzgün cavab",
        "Ümumi sual",
        "Səhv cavab",
        "Nəticə",
        "Tarix"
    ]

    for col, header in enumerate(
        headers,
        1
    ):

        cell = sheet.cell(
            row=1,
            column=col,
            value=header
        )

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # NƏTİCƏLƏR
    for row_number, r in enumerate(
        results,
        2
    ):

        sheet.cell(
            row=row_number,
            column=1,
            value=row_number - 1
        )

        sheet.cell(
            row=row_number,
            column=2,
            value=r["name"]
        )

        sheet.cell(
            row=row_number,
            column=3,
            value=r["correct"]
        )

        sheet.cell(
            row=row_number,
            column=4,
            value=r["total"]
        )

        sheet.cell(
            row=row_number,
            column=5,
            value=r["total"] - r["correct"]
        )

        sheet.cell(
            row=row_number,
            column=6,
            value=f'{r["percent"]}%'
        )

        sheet.cell(
            row=row_number,
            column=7,
            value=r["created_at"].strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            if r["created_at"]
            else ""
        )

    # SÜTUN ENLİKLƏRİ
    widths = {
        "A": 8,
        "B": 30,
        "C": 18,
        "D": 18,
        "E": 15,
        "F": 15,
        "G": 25
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width

    # FAYLI YADDAŞDA YARAT
    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    filename = (
        "emek_mecellesi_2026_test_neticeleri.xlsx"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ==========================================================
# RUN
# ==========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        )
    )
