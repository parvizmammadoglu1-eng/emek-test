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
# SUALLAR
# =========================================================

QUESTIONS = [

    (
        "Əmək Məcəlləsi kimlərə şamil edilir?",
        "a) əcnəbilərə",
        "b) hərbi qulluqçulara",
        "c) məhkəmə hakimlərinə",
        "d) AR-nın Milli Məclisinin deputatlarına və bələdiyyələrə seçilmiş şəxslərə",
        "A"
    ),

    (
        "Əmək qanunvericiliyinə əməl olunmasına dövlət nəzarətini hansı orqan həyata keçirir?",
        "a) rayon (şəhər) məhkəməsi",
        "b) rayon (şəhər) məşğulluq mərkəzləri",
        "c) Azərbaycan Həmkarlar İttifaqları Konfederasiyası",
        "d) Dövlət Əmək Müfəttişliyi",
        "D"
    ),

    (
        "Əmək müqaviləsinin tərəfləri kimlər olur?",
        "a) işçi və işəgötürən",
        "b) işçi və həmkarlar ittifaqı təşkilatı",
        "c) işçi və əmək kollektivi",
        "d) işəgötürən və həmkarlar ittifaqı təşkilatı",
        "A"
    ),

    (
        "Hansı yaşdan hər bir şəxs işçi kimi əmək müqaviləsinin tərəfi ola bilər?",
        "a) 13 yaşdan",
        "b) 14 yaşdan",
        "c) 15 yaşdan",
        "d) 16 yaşdan",
        "C"
    ),

    (
        "Əmək münasibətlərini hansı hüquqi fakt yaradır?",
        "a) kollektiv müqavilə",
        "b) əmək müqaviləsi",
        "c) mülki-hüquqi müqavilə",
        "d) işəgötürənin əmri (sərəncamı, qərarı)",
        "B"
    ),

    (
        "Ezamiyyətin müddəti neçə gündən artıq ola bilməz?",
        "a) 30 təqvim günündən",
        "b) 40 təqvim günündən",
        "c) 45 təqvim günündən",
        "d) 25 təqvim günündən",
        "A"
    ),

    (
        "İşçiyə məzuniyyət vaxtı üçün orta əmək haqqı məzuniyyətin başlanmasına ən azı neçə gün qalmış ödənilir?",
        "a) 3 gün qalmış",
        "b) 4 gün qalmış",
        "c) 5 gün qalmış",
        "d) 6 gün qalmış",
        "A"
    ),

    (
        "İşçinin on ildən on beş ilədək əmək stajı olduqda əlavə neçə gün məzuniyyət verilir?",
        "a) 8 təqvim günü",
        "b) 5 təqvim günü",
        "c) 6 təqvim günü",
        "d) 4 təqvim günü",
        "B"
    ),

    (
        "İşçinin bir iş günü ilə növbəti iş günü arasındakı gündəlik istirahət vaxtı azı neçə saat olmalıdır?",
        "a) azı 8 saat",
        "b) azı 10 saat",
        "c) azı 12 saat",
        "d) azı 14 saat",
        "C"
    ),

    (
        "16 yaşdan 18 yaşadək olan işçilərə qısaldılmış iş vaxtının müddəti həftə ərzində neçə saat təşkil edir?",
        "a) 24 saat",
        "b) 36 saat",
        "c) 40 saat",
        "d) 32 saat",
        "B"
    )

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


init_db()


# =========================================================
# ADMIN YOXLANIŞI
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
# ANA SƏHİFƏ
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


# =========================================================
# TEST SUALLARI
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

    con = db()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
    ).fetchall()

    con.close()

    total = len(questions)

    if total == 0:
        return redirect(
            url_for("home")
        )

    if n >= total:
        return redirect(
            url_for("finish")
        )

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        answers = session.get(
            "answers",
            {}
        )

        # Cavab verilibsə yadda saxla
        if selected in ["A", "B", "C", "D"]:

            answers[str(n)] = selected

            session["answers"] = answers

        # TESTİ BİTİR düyməsi
        action = request.form.get(
            "action"
        )

        if action == "finish":

            return redirect(
                url_for("finish")
            )

        # NÖVBƏTİ
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
        total=total,
        name=session["name"]
    )


# =========================================================
# NƏTİCƏ
# =========================================================

@app.route("/finish")
def finish():

    if "name" not in session:
        return redirect(
            url_for("home")
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

    # Yalnız cavablandırılmış suallar
    answered_count = len(answers)

    # Düzgün cavab sayı
    correct = 0

    for i, q in enumerate(questions):

        selected = answers.get(
            str(i)
        )

        if selected == q["answer"]:
            correct += 1

    # İştirakçı heç bir suala cavab verməyibsə
    if answered_count > 0:

        percent = round(
            correct / answered_count * 100
        )

    else:

        percent = 0

    name = session["name"]

    # Nəticəni yadda saxla
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
            answered_count,
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
        total=answered_count,
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
        "SELECT * FROM results ORDER BY id DESC"
    ).fetchall()

    questions = con.execute(
        "SELECT * FROM questions ORDER BY id"
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
        "SELECT * FROM results ORDER BY id DESC"
    ).fetchall()

    con.close()

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Test nəticələri"

    headers = [
        "№",
        "Ad və soyad",
        "Düzgün cavab",
        "Cavablandırılan sual",
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
            value=r["created_at"]
        )

    widths = {
        "A": 8,
        "B": 30,
        "C": 18,
        "D": 24,
        "E": 15,
        "F": 15,
        "G": 25
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width

    output = BytesIO()

    workbook.save(
        output
    )

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
