import os
from datetime import datetime
from zoneinfo import ZoneInfo
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

import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# FLASK
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

GOOGLE_SHEET_NAME = "Emek Test 2026"


def get_google_client():

    private_key = os.environ.get(
        "GOOGLE_PRIVATE_KEY",
        ""
    )

    if not private_key:
        raise RuntimeError(
            "GOOGLE_PRIVATE_KEY Render Environment Variables bölməsində yoxdur."
        )

    # Render-də \n mətn kimi gəlirsə,
    # onu real yeni sətrə çeviririk.
    private_key = private_key.replace(
        "\\n",
        "\n"
    )

    service_account_info = {
        "type": "service_account",

        "project_id": os.environ.get(
            "GOOGLE_PROJECT_ID"
        ),

        "private_key_id": os.environ.get(
            "GOOGLE_PRIVATE_KEY_ID"
        ),

        "private_key": private_key,

        "client_email": os.environ.get(
            "GOOGLE_CLIENT_EMAIL"
        ),

        "token_uri":
            "https://oauth2.googleapis.com/token"
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=scopes
    )

    return gspread.authorize(
        credentials
    )


def get_sheet():

    client = get_google_client()

    spreadsheet = client.open(
        GOOGLE_SHEET_NAME
    )

    try:

        sheet = spreadsheet.worksheet(
            "Nəticələr"
        )

    except gspread.WorksheetNotFound:

        sheet = spreadsheet.add_worksheet(
            title="Nəticələr",
            rows=1000,
            cols=8
        )

        sheet.append_row(
            [
                "№",
                "Ad və soyad",
                "Düzgün cavab",
                "Ümumi sual",
                "Səhv cavab",
                "Nəticə",
                "Tarix",
                "Status"
            ]
        )

    return sheet


# =========================================================
# SUALLAR
# =========================================================

QUESTIONS = [

    {
        "question":
            "Əmək Məcəlləsi kimlərə şamil edilir?",

        "a":
            "a) əcnəbilərə;",

        "b":
            "b) hərbi qulluqçulara;",

        "c":
            "c) məhkəmə hakimlərinə;",

        "d":
            "d) AR-nın Milli Məclisinin deputatlarına və bələdiyyələrə seçilmiş şəxslərə;",

        "answer":
            "A"
    },

    {
        "question":
            "Əmək qanunvericiliyinə əməl olunmasına dövlət nəzarətini hansı orqan həyata keçirir?",

        "a":
            "a) rayon (şəhər) məhkəməsi;",

        "b":
            "b) rayon (şəhər) məşğulluq mərkəzləri;",

        "c":
            "c) Azərbaycan Həmkarlar İttifaqları Konfederasiyası;",

        "d":
            "d) Dövlət Əmək Müfəttişliyi;",

        "answer":
            "D"
    },

    {
        "question":
            "Əmək müqaviləsinin tərəfləri kimlər olur?",

        "a":
            "a) işçi və işəgötürən;",

        "b":
            "b) işçi və həmkarlar ittifaqı təşkilatı;",

        "c":
            "c) işçi və əmək kollektivi;",

        "d":
            "d) işəgötürən və həmkarlar ittifaqı təşkilatı",

        "answer":
            "A"
    },

    {
        "question":
            "Hansı yaşdan hər bir şəxs işçi kimi əmək müqaviləsinin tərəfi ola bilər?",

        "a":
            "a) 13 yaşdan;",

        "b":
            "b) 14 yaşdan;",

        "c":
            "c) 15 yaşdan;",

        "d":
            "d) 16 yaşdan",

        "answer":
            "C"
    },

    {
        "question":
            "Əmək münasibətlərini hansı hüquqi fakt yaradır?",

        "a":
            "a) kollektiv müqavilə;",

        "b":
            "b) əmək müqaviləsi;",

        "c":
            "c) mülki-hüquqi müqavilə;",

        "d":
            "d) işəgötürənin əmri (sərəncamı, qərarı)",

        "answer":
            "B"
    },

    {
        "question":
            "Ezamiyyətin müddəti neçə gündən artıq ola bilməz?",

        "a":
            "a) 30 təqvim günündən",

        "b":
            "b) 40 təqvim günündən",

        "c":
            "c) 45 təqvim günündən",

        "d":
            "d) 25 təqvim günündən",

        "answer":
            "A"
    },

    {
        "question":
            "İşçiyə məzuniyyət vaxtı üçün orta əmək haqqı məzuniyyətin başlanmasına ən azı neçə gün qalmış ödənilir?",

        "a":
            "a) 3 gün qalmış",

        "b":
            "b) 4 gün qalmış",

        "c":
            "c) 5 gün qalmış",

        "d":
            "d) 6 gün qalmış",

        "answer":
            "C"
    },

    {
        "question":
            "İşçinin on ildən on beş ilədək əmək stajı olduqda əlavə necə gün məzuniyyət verilir?",

        "a":
            "a) 8 təqvim günü",

        "b":
            "b) 5 təqvim günü",

        "c":
            "c) 6 təqvim günü",

        "d":
            "d) 4 təqvim günü",

        "answer":
            "B"
    },

    {
        "question":
            "İşçinin bir iş günü ilə növbəti iş günü arasındakı gündəlik istirahət vaxtı azı neçə saat olmalıdır?",

        "a":
            "a) azı 8 saat",

        "b":
            "b) azı 10 saat",

        "c":
            "c) azı 12 saat",

        "d":
            "d) azı 14 saat",

        "answer":
            "C"
    },

    {
        "question":
            "16 yaşdan 18 yaşadək olan işçilərə qısaldılmış iş vaxtının müddəti həftə ərzində neçə saat təşkil edir?",

        "a":
            "a) 24 saat",

        "b":
            "b) 36 saat",

        "c":
            "c) 40 saat",

        "d":
            "d) 32 saat",

        "answer":
            "B"
    }

]


# =========================================================
# ADMIN
# =========================================================

def admin_required(function):

    def wrapper(*args, **kwargs):

        if not session.get("admin"):

            return redirect(
                url_for("admin_login")
            )

        return function(
            *args,
            **kwargs
        )

    wrapper.__name__ = function.__name__

    return wrapper


# =========================================================
# HOME
# =========================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
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

    total = len(
        QUESTIONS
    )

    # Sual sayı bitibsə nəticəyə keç
    if n >= total:

        return redirect(
            url_for("finish")
        )

    q = QUESTIONS[n]

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        answers = session.get(
            "answers",
            {}
        )

        # =================================================
        # CAVAB SEÇİLİBSƏ YADDA SAXLA
        # =================================================

        if selected in [
            "A",
            "B",
            "C",
            "D"
        ]:

            answers[str(n)] = selected

        # =================================================
        # CAVAB SEÇİLMƏYİBSƏ BOŞ SAXLA
        # =================================================

        else:

            answers.pop(
                str(n),
                None
            )

        session["answers"] = answers

        session.modified = True

        # =================================================
        # SON SUALDIRSA NƏTİCƏYƏ KEÇ
        # =================================================

        if n + 1 >= total:

            return redirect(
                url_for("finish")
            )

        # =================================================
        # NÖVBƏTİ SUALA KEÇ
        # =================================================

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

@app.route(
    "/finish",
    methods=["GET", "POST"]
)
def finish():

    if "name" not in session:

        return redirect(
            url_for("home")
        )

    answers = session.get(
        "answers",
        {}
    )

    total = len(
        QUESTIONS
    )

    correct = 0

    answered = len(
        answers
    )

    # =====================================================
    # SUALLAR ÜZRƏ ƏTRAFLI NƏTİCƏ HAZIRLA
    # =====================================================

    question_results = []

    for index, q in enumerate(QUESTIONS):

        selected = answers.get(
            str(index)
        )

        correct_letter = q["answer"]

        # -------------------------------------------------
        # DÜZGÜN CAVABIN TAM MƏTNİ
        # -------------------------------------------------

        correct_text = q[
            correct_letter.lower()
        ]

        # -------------------------------------------------
        # İŞTİRAKÇININ SEÇDİYİ CAVABIN TAM MƏTNİ
        # -------------------------------------------------

        if selected in [
            "A",
            "B",
            "C",
            "D"
        ]:

            selected_text = q[
                selected.lower()
            ]

        else:

            selected_text = "Cavab verilməyib"

        # -------------------------------------------------
        # NƏTİCƏNİ MÜƏYYƏN ET
        # -------------------------------------------------

        if selected == correct_letter:

            is_correct = True
            is_wrong = False
            is_empty = False

            correct += 1

        elif selected is None:

            is_correct = False
            is_wrong = False
            is_empty = True

        else:

            is_correct = False
            is_wrong = True
            is_empty = False

        # -------------------------------------------------
        # HTML-Ə GÖNDƏRİLƏCƏK MƏLUMATLAR
        # -------------------------------------------------

        question_results.append({

            "number":
                index + 1,

            "question":
                q["question"],

            "selected":
                selected,

            "selected_text":
                selected_text,

            "correct_answer":
                correct_letter,

            "correct_text":
                correct_text,

            "is_correct":
                is_correct,

            "is_wrong":
                is_wrong,

            "is_empty":
                is_empty

        })

    # =====================================================
    # SƏHV CAVABLAR
    # =====================================================

    wrong = answered - correct

    # =====================================================
    # BOŞ CAVABLAR
    # =====================================================

    unanswered = total - answered

    # =====================================================
    # FAİZ
    # =====================================================

    percent = round(
        correct / total * 100
    ) if total else 0

    name = session["name"]

    # =====================================================
    # BAKI VAXTI
    # =====================================================

    created_at = datetime.now(
        ZoneInfo("Asia/Baku")
    ).strftime(
        "%d.%m.%Y %H:%M:%S"
    )

    # =====================================================
    # STATUS
    # =====================================================

    if answered == total:

        status = "Tamamlandı"

    else:

        status = (
            f"Yarımçıq bitirildi "
            f"({answered}/{total})"
        )

    # =====================================================
    # GOOGLE SHEETS
    # =====================================================

    try:

        sheet = get_sheet()

        all_values = sheet.get_all_values()

        # Başlıq sətrini nəzərə al
        number = len(
            all_values
        )

        sheet.append_row(
            [
                number,
                name,
                correct,
                total,
                wrong,
                f"{percent}%",
                created_at,
                status
            ],
            value_input_option="USER_ENTERED"
        )

        print(
            "GOOGLE SHEETS: nəticə əlavə edildi."
        )

    except Exception as e:

        print(
            "GOOGLE SHEETS ERROR:",
            str(e)
        )

    # =====================================================
    # VACİB:
    #
    # SESSION-I BURADA TƏMİZLƏMİRİK.
    #
    # Çünki finish.html nəticələri göstərmək üçün
    # məlumatdan istifadə edir.
    # =====================================================

    return render_template(

        "finish.html",

        name=name,

        correct=correct,

        total=total,

        percent=percent,

        answered=answered,

        wrong=wrong,

        unanswered=unanswered,

        status=status,

        question_results=question_results

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

@app.route(
    "/admin/logout"
)
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

@app.route(
    "/admin"
)
@admin_required
def admin():

    try:

        sheet = get_sheet()

        values = sheet.get_all_records()

    except Exception as e:

        print(
            "ADMIN GOOGLE SHEETS ERROR:",
            str(e)
        )

        values = []

    return render_template(
        "admin.html",

        results=values,

        questions=QUESTIONS
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route(
    "/admin/export"
)
@admin_required
def admin_export():

    try:

        sheet = get_sheet()

        results = sheet.get_all_records()

    except Exception as e:

        print(
            "EXPORT GOOGLE SHEETS ERROR:",
            str(e)
        )

        results = []

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Test nəticələri"

    # =====================================================
    # BAŞLIQLAR
    # =====================================================

    headers = [

        "№",

        "Ad və soyad",

        "Düzgün cavab",

        "Ümumi sual",

        "Səhv cavab",

        "Nəticə",

        "Tarix",

        "Status"

    ]

    for column, header in enumerate(
        headers,
        1
    ):

        cell = worksheet.cell(
            row=1,
            column=column,
            value=header
        )

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # =====================================================
    # NƏTİCƏLƏR
    # =====================================================

    for row_number, result in enumerate(
        results,
        2
    ):

        worksheet.cell(
            row=row_number,
            column=1,
            value=result.get(
                "№",
                row_number - 1
            )
        )

        worksheet.cell(
            row=row_number,
            column=2,
            value=result.get(
                "Ad və soyad",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=3,
            value=result.get(
                "Düzgün cavab",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=4,
            value=result.get(
                "Ümumi sual",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=5,
            value=result.get(
                "Səhv cavab",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=6,
            value=result.get(
                "Nəticə",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=7,
            value=result.get(
                "Tarix",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=8,
            value=result.get(
                "Status",
                ""
            )
        )

    # =====================================================
    # SÜTUN ENLİKLƏRİ
    # =====================================================

    widths = {

        "A": 8,

        "B": 30,

        "C": 18,

        "D": 18,

        "E": 15,

        "F": 15,

        "G": 25,

        "H": 30

    }

    for column, width in widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    # =====================================================
    # EXCEL FAYLI
    # =====================================================

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
