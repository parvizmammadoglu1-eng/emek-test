import os
import json
from datetime import datetime
from functools import wraps
from io import BytesIO

import requests
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


# =========================================================
# APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "temporary-secret-key-change-me"
)

ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD",
    "12345"
)


# =========================================================
# GOOGLE SHEETS
# =========================================================

GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_JSON",
    ""
)


def google_sheets_available():
    return bool(
        GOOGLE_SHEET_ID and
        GOOGLE_SERVICE_ACCOUNT_JSON
    )


def get_google_credentials():
    """
    Google Service Account məlumatlarını oxuyur.
    """
    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        return None

    try:
        from google.oauth2.service_account import Credentials

        data = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        credentials = Credentials.from_service_account_info(
            data,
            scopes=scopes
        )

        return credentials

    except Exception as e:
        print("Google credentials error:", e)
        return None


def save_result_to_google_sheet(
    name,
    correct,
    total,
    percent,
    created_at
):
    """
    Nəticəni Google Sheets-ə əlavə edir.
    """

    if not google_sheets_available():
        print("Google Sheets environment variables yoxdur.")
        return False

    try:

        import gspread

        credentials = get_google_credentials()

        if credentials is None:
            return False

        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_key(
            GOOGLE_SHEET_ID
        )

        worksheet = spreadsheet.sheet1

        wrong = total - correct

        worksheet.append_row(
            [
                name,
                correct,
                total,
                wrong,
                percent,
                created_at
            ],
            value_input_option="USER_ENTERED"
        )

        return True

    except Exception as e:

        print("Google Sheets save error:", e)

        return False


def get_results_from_google_sheet():
    """
    Google Sheets-dən bütün nəticələri oxuyur.
    """

    if not google_sheets_available():
        return []

    try:

        import gspread

        credentials = get_google_credentials()

        if credentials is None:
            return []

        client = gspread.authorize(credentials)

        spreadsheet = client.open_by_key(
            GOOGLE_SHEET_ID
        )

        worksheet = spreadsheet.sheet1

        rows = worksheet.get_all_records()

        results = []

        for index, row in enumerate(rows, 1):

            results.append(
                {
                    "id": index,
                    "name": row.get("Ad və soyad", ""),
                    "correct": int(
                        row.get("Düzgün cavab", 0) or 0
                    ),
                    "total": int(
                        row.get("Ümumi sual", 0) or 0
                    ),
                    "percent": int(
                        row.get("Nəticə", 0) or 0
                    ),
                    "created_at": row.get(
                        "Tarix",
                        ""
                    )
                }
            )

        results.reverse()

        return results

    except Exception as e:

        print("Google Sheets read error:", e)

        return []


# =========================================================
# QUESTIONS
# =========================================================

QUESTIONS = [

    {
        "id": 1,
        "question": "Əmək Məcəlləsi kimlərə şamil edilir?",
        "a": "a) əcnəbilərə;",
        "b": "b) hərbi qulluqçulara;",
        "c": "c) məhkəmə hakimlərinə;",
        "d": "d) AR-nın Milli Məclisinin deputatlarına və bələdiyyələrə seçilmiş şəxslərə;",
        "answer": "A"
    },

    {
        "id": 2,
        "question": "Əmək qanunvericiliyinə əməl olunmasına dövlət nəzarətini hansı orqan həyata keçirir?",
        "a": "a) rayon (şəhər) məhkəməsi;",
        "b": "b) rayon (şəhər) məşğulluq mərkəzləri;",
        "c": "c) Azərbaycan Həmkarlar İttifaqları Konfederasiyası;",
        "d": "d) Dövlət Əmək Müfəttişliyi;",
        "answer": "D"
    },

    {
        "id": 3,
        "question": "Əmək müqaviləsinin tərəfləri kimlər olur?",
        "a": "a) işçi və işəgötürən;",
        "b": "b) işçi və həmkarlar ittifaqı təşkilatı;",
        "c": "c) işçi və əmək kollektivi;",
        "d": "d) işəgötürən və həmkarlar ittifaqı təşkilatı",
        "answer": "A"
    },

    {
        "id": 4,
        "question": "Hansı yaşdan hər bir şəxs işçi kimi əmək müqaviləsinin tərəfi ola bilər?",
        "a": "a) 13 yaşdan;",
        "b": "b) 14 yaşdan;",
        "c": "c) 15 yaşdan;",
        "d": "d) 16 yaşdan",
        "answer": "D"
    },

    {
        "id": 5,
        "question": "Əmək münasibətlərini hansı hüquqi fakt yaradır?",
        "a": "a) kollektiv müqavilə;",
        "b": "b) əmək müqaviləsi;",
        "c": "c) mülki-hüquqi müqavilə;",
        "d": "d) işəgötürənin əmri (sərəncamı, qərarı)",
        "answer": "B"
    },

    {
        "id": 6,
        "question": "Ezamiyyətin müddəti neçə gündən artıq ola bilməz?",
        "a": "a) 30 təqvim günündən",
        "b": "b) 40 təqvim günündən",
        "c": "c) 45 təqvim günündən",
        "d": "d) 25 təqvim günündən",
        "answer": "A"
    },

    {
        "id": 7,
        "question": "İşçiyə məzuniyyət vaxtı üçün orta əmək haqqı məzuniyyətin başlanmasına ən azı neçə gün qalmış ödənilir?",
        "a": "a) 3 gün qalmış",
        "b": "b) 4 gün qalmış",
        "c": "c) 5 gün qalmış",
        "d": "d) 6 gün qalmış",
        "answer": "C"
    },

    {
        "id": 8,
        "question": "İşçinin on ildən on beş ilədək əmək stajı olduqda əlavə necə gün məzuniyyət verilir?",
        "a": "a) 8 təqvim günü",
        "b": "b) 5 təqvim günü",
        "c": "c) 6 təqvim günü",
        "d": "d) 4 təqvim günü",
        "answer": "B"
    },

    {
        "id": 9,
        "question": "İşçinin bir iş günü ilə növbəti iş günü arasındakı gündəlik istirahət vaxtı azı neçə saat olmalıdır?",
        "a": "a) azı 8 saat",
        "b": "b) azı 10 saat",
        "c": "c) azı 12 saat",
        "d": "d) azı 14 saat",
        "answer": "C"
    },

    {
        "id": 10,
        "question": "16 yaşdan 18 yaşadək olan işçilərə qısaldılmış iş vaxtının müddəti həftə ərzində neçə saat təşkil edir?",
        "a": "a) 24 saat",
        "b": "b) 36 saat",
        "c": "c) 40 saat",
        "d": "d) 32 saat",
        "answer": "B"
    }

]


# =========================================================
# ADMIN
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
# HOME
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
# QUESTION
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

    total = len(QUESTIONS)

    if n >= total:
        return redirect(
            url_for("finish")
        )

    q = QUESTIONS[n]

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        finish_now = request.form.get(
            "finish_now"
        )

        if selected not in [
            "A",
            "B",
            "C",
            "D"
        ]:

            return render_template(
                "question.html",
                q=q,
                n=n,
                total=total,
                name=session["name"],
                error="Zəhmət olmasa cavablardan birini seçin."
            )

        answers = session.get(
            "answers",
            {}
        )

        answers[str(n)] = selected

        session["answers"] = answers

        # İştirakçı "TESTİ BİTİR" düyməsini basıbsa
        if finish_now:

            return redirect(
                url_for("finish")
            )

        return redirect(
            url_for(
                "question",
                n=n + 1
            )
        )

    return render_template(
        "question.html",
        q=q,
        n=n,
        total=total,
        name=session["name"]
    )


# =========================================================
# FINISH
# =========================================================

@app.route("/finish")
def finish():

    if "name" not in session:

        return redirect(
            url_for("home")
        )

    answers = session.get(
        "answers",
        {}
    )

    correct = 0

    answered = 0

    for i, q in enumerate(QUESTIONS):

        selected = answers.get(
            str(i)
        )

        if selected:

            answered += 1

            if selected == q["answer"]:

                correct += 1

    total = len(QUESTIONS)

    percent = round(
        correct / total * 100
    ) if total else 0

    name = session["name"]

    created_at = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    # GOOGLE SHEETS-Ə YAZ
    save_result_to_google_sheet(
        name=name,
        correct=correct,
        total=total,
        percent=percent,
        created_at=created_at
    )

    session.clear()

    return render_template(
        "finish.html",
        name=name,
        correct=correct,
        total=total,
        percent=percent,
        answered=answered
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

    results = get_results_from_google_sheet()

    questions = QUESTIONS

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

    results = get_results_from_google_sheet()

    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Test nəticələri"

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
            value=(
                r["total"] -
                r["correct"]
            )
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
        "D": 18,
        "E": 15,
        "F": 15,
        "G": 25
    }

    for column, width in widths.items():

        sheet.column_dimensions[
            column
        ].width = width

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
