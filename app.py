import os
import sqlite3
from datetime import datetime
from pathlib import Path
from functools import wraps
from io import BytesIO

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

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "temporary-secret-key-change-me"
)

ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "12345"
)

DB = Path(__file__).with_name("test.db")


# =========================================================
# İLK SUALLAR
# =========================================================

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


# =========================================================
# DATABASE
# =========================================================

def db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con


def init_db():
    con = db()

    # Suallar
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

    # Nəticələr
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

    # -----------------------------------------------------
    # Mövcud DB-də yeni sütunlar yoxdursa əlavə edirik
    # -----------------------------------------------------

    columns = [
        row["name"]
        for row in con.execute(
            "PRAGMA table_info(results)"
        ).fetchall()
    ]

    if "answered" not in columns:
        con.execute(
            "ALTER TABLE results "
            "ADD COLUMN answered INTEGER NOT NULL DEFAULT 0"
        )

    if "wrong" not in columns:
        con.execute(
            "ALTER TABLE results "
            "ADD COLUMN wrong INTEGER NOT NULL DEFAULT 0"
        )

    # -----------------------------------------------------
    # Əgər sual bazası boşdursa ilkin sualları əlavə et
    # -----------------------------------------------------

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


init_db()


# =========================================================
# ADMIN DECORATOR
# =========================================================

def admin_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("admin"):
            return redirect(
                url_for("admin_login")
            )

        return function(*args, **kwargs)

    return wrapper


# =========================================================
# İSTİFADƏÇİ
# =========================================================

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

        # Yeni test başlayır
        session["name"] = name
        session["answers"] = {}

        # Test artıq tamamlanmayıb
        session["test_finished"] = False

        return redirect(
            url_for(
                "question",
                n=0
            )
        )

    return render_template("home.html")


# =========================================================
# SUAL
# =========================================================

@app.route(
    "/question/<int:n>",
    methods=["GET", "POST"]
)
def question(n):

    if "name" not in session:
        return redirect(
            url_for("home")
        )

    # Əgər test artıq bitibsə
    if session.get("test_finished"):
        return redirect(
            url_for("finish")
        )

    con = db()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
    ).fetchall()

    con.close()

    # Sual sayı bitibsə avtomatik nəticəyə keç
    if n >= len(questions):
        return redirect(
            url_for("finish")
        )

    # -----------------------------------------------------
    # NÖVBƏTİ düyməsi
    # -----------------------------------------------------

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        # Cavab seçilməyibsə
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

        # Cavabı yadda saxla
        answers[str(n)] = selected

        session["answers"] = answers
        session.modified = True

        # Növbəti suala keç
        return redirect(
            url_for(
                "question",
                n=n + 1
            )
        )

    # Hazırkı cavab
    answers = session.get(
        "answers",
        {}
    )

    current_answer = answers.get(
        str(n)
    )

    return render_template(
        "question.html",
        q=questions[n],
        n=n,
        total=len(questions),
        name=session["name"],
        current_answer=current_answer
    )


# =========================================================
# TESTİ BİTİR
# =========================================================

@app.route("/finish", methods=["GET", "POST"])
def finish():

    if "name" not in session:
        return redirect(
            url_for("home")
        )

    # Əgər nəticə artıq hesablanıbsa,
    # yenidən DB-yə yazma
    if session.get("test_finished"):

        return render_template(
            "finish.html",
            name=session.get("finished_name", ""),
            correct=session.get("finished_correct", 0),
            wrong=session.get("finished_wrong", 0),
            answered=session.get("finished_answered", 0),
            total=session.get("finished_total", 0),
            percent=session.get("finished_percent", 0)
        )

    con = db()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
    ).fetchall()

    con.close()

    answers = session.get(
        "answers",
        {}
    )

    # -----------------------------------------------------
    # Nəticənin hesablanması
    # -----------------------------------------------------

    correct = 0
    answered = 0

    for i, q in enumerate(questions):

        selected = answers.get(
            str(i)
        )

        # Cavablandırılmayıbsa heç bir
        # kateqoriyaya daxil edilmir
        if selected is None:
            continue

        answered += 1

        if selected == q["answer"]:
            correct += 1

    total = len(questions)

    # Cavablandırılan sualların içindən
    # düzgün cavabların faizi
    if answered > 0:
        percent = round(
            correct / answered * 100
        )
    else:
        percent = 0

    wrong = answered - correct

    name = session["name"]

    # -----------------------------------------------------
    # Nəticəni database-ə yaz
    # -----------------------------------------------------

    con = db()

    con.execute(
        """
        INSERT INTO results
        (
            name,
            correct,
            total,
            percent,
            created_at,
            answered,
            wrong
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            name,
            correct,
            total,
            percent,
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            answered,
            wrong
        )
    )

    con.commit()
    con.close()

    # -----------------------------------------------------
    # Nəticəni session-da saxla
    # -----------------------------------------------------

    session["test_finished"] = True

    session["finished_name"] = name
    session["finished_correct"] = correct
    session["finished_wrong"] = wrong
    session["finished_answered"] = answered
    session["finished_total"] = total
    session["finished_percent"] = percent

    # Testin aktiv cavablarını artıq lazım deyil
    session.pop("answers", None)

    return render_template(
        "finish.html",
        name=name,
        correct=correct,
        wrong=wrong,
        answered=answered,
        total=total,
        percent=percent
    )


# =========================================================
# ADMIN LOGIN
# =========================================================

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


# =========================================================
# ADMIN LOGOUT
# =========================================================

@app.route("/admin/logout")
def admin_logout():

    session.pop(
        "admin",
        None
    )

    return redirect(
        url_for("admin_login")
    )


# =========================================================
# ADMIN PANEL
# =========================================================

@app.route("/admin")
@admin_required
def admin():

    con = db()

    results = con.execute(
        """
        SELECT *
        FROM results
        ORDER BY id DESC
        """
    ).fetchall()

    questions = con.execute(
        """
        SELECT *
        FROM questions
        ORDER BY id
        """
    ).fetchall()

    con.close()

    return render_template(
        "admin.html",
        results=results,
        questions=questions
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route("/admin/export")
@admin_required
def admin_export():

    con = db()

    results = con.execute(
        """
        SELECT *
        FROM results
        ORDER BY id DESC
        """
    ).fetchall()

    con.close()

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Test nəticələri"

    # -----------------------------------------------------
    # Başlıqlar
    # -----------------------------------------------------

    headers = [
        "№",
        "Ad və soyad",
        "Cavablandırılan",
        "Düzgün cavab",
        "Səhv cavab",
        "Ümumi sual",
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

    # -----------------------------------------------------
    # Nəticələr
    # -----------------------------------------------------

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
            value=r["answered"]
        )

        sheet.cell(
            row=row_number,
            column=4,
            value=r["correct"]
        )

        sheet.cell(
            row=row_number,
            column=5,
            value=r["wrong"]
        )

        sheet.cell(
            row=row_number,
            column=6,
            value=r["total"]
        )

        sheet.cell(
            row=row_number,
            column=7,
            value=f'{r["percent"]}%'
        )

        sheet.cell(
            row=row_number,
            column=8,
            value=r["created_at"]
        )

    # -----------------------------------------------------
    # Sütun genişlikləri
    # -----------------------------------------------------

    widths = {
        "A": 8,
        "B": 30,
        "C": 18,
        "D": 18,
        "E": 15,
        "F": 15,
        "G": 15,
        "H": 25
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width

    # -----------------------------------------------------
    # Excel faylı
    # -----------------------------------------------------

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


# =========================================================
# RUN
# =========================================================

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
