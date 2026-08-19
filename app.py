import os
import json
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

from openpyxl import Workbook, load_workbook
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


# =========================================================
# NƏTİCƏLƏR SHEET
# =========================================================

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
            cols=11
        )

    headers = sheet.row_values(1)

    required_headers = [
        "№",
        "Ad və soyad",
        "Düzgün cavab",
        "Ümumi sual",
        "Səhv cavab",
        "Nəticə",
        "Tarix",
        "Status",
        "Müddət",
        "Başlama vaxtı",
        "Bitmə vaxtı"
    ]

    if not headers:

        sheet.update(
            "1:1",
            [required_headers]
        )

    else:

        changed = False

        for header in required_headers:

            if header not in headers:

                headers.append(
                    header
                )

                changed = True

        if changed:

            sheet.resize(
                rows=max(
                    sheet.row_count,
                    1000
                ),
                cols=max(
                    sheet.col_count,
                    len(headers)
                )
            )

            sheet.update(
                "1:1",
                [headers]
            )

    return sheet


# =========================================================
# SUALLAR SHEET
# =========================================================

QUESTIONS_SHEET_NAME = "Suallar"

QUESTIONS_HEADERS = [
    "question",
    "a",
    "b",
    "c",
    "d",
    "answer"
]


def get_questions_sheet():

    client = get_google_client()

    spreadsheet = client.open(
        GOOGLE_SHEET_NAME
    )

    try:

        sheet = spreadsheet.worksheet(
            QUESTIONS_SHEET_NAME
        )

    except gspread.WorksheetNotFound:

        sheet = spreadsheet.add_worksheet(
            title=QUESTIONS_SHEET_NAME,
            rows=1000,
            cols=6
        )

        sheet.update(
            "A1:F1",
            [QUESTIONS_HEADERS]
        )

    return sheet


# =========================================================
# SUALLARI GOOGLE SHEETS-DƏN OXU
# =========================================================

def load_questions():

    try:

        sheet = get_questions_sheet()

        values = sheet.get_all_values()

        if len(values) <= 1:

            return []

        questions = []

        for row in values[1:]:

            if not row:
                continue

            if len(row) < 6:
                continue

            question_text = str(
                row[0]
            ).strip()

            if not question_text:
                continue

            answer = str(
                row[5]
            ).strip().upper()

            if answer not in [
                "A",
                "B",
                "C",
                "D"
            ]:
                continue

            questions.append({

                "question":
                    question_text,

                "a":
                    str(row[1]).strip(),

                "b":
                    str(row[2]).strip(),

                "c":
                    str(row[3]).strip(),

                "d":
                    str(row[4]).strip(),

                "answer":
                    answer

            })

        return questions

    except Exception as e:

        print(
            "LOAD QUESTIONS ERROR:",
            str(e)
        )

        return []


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

        session.clear()

        session["name"] = name

        session["answers"] = {}

        start_time = datetime.now(
            ZoneInfo("Asia/Baku")
        )

        session["exam_start_time"] = (
            start_time.timestamp()
        )

        session.modified = True

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

    current_questions = load_questions()

    total = len(
        current_questions
    )

    if total == 0:

        return render_template(
            "home.html",
            error="Hazırda sistemdə sual yoxdur."
        )

    if n >= total:

        return redirect(
            url_for("finish")
        )

    q = current_questions[n]

    if request.method == "POST":

        selected = request.form.get(
            "answer"
        )

        answers = session.get(
            "answers",
            {}
        )

        if selected in [
            "A",
            "B",
            "C",
            "D"
        ]:

            answers[str(n)] = selected

        else:

            answers.pop(
                str(n),
                None
            )

        session["answers"] = answers

        session.modified = True

        if n + 1 >= total:

            return redirect(
                url_for("finish")
            )

        return redirect(
            url_for(
                "question",
                n=n + 1
            )
        )

    exam_start_time = session.get(
        "exam_start_time"
    )

    return render_template(

        "question.html",

        q=q,

        n=n,

        total=total,

        name=session["name"],

        exam_start_time=exam_start_time

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

    current_questions = load_questions()

    total = len(
        current_questions
    )

    if total == 0:

        return redirect(
            url_for("home")
        )

    correct = 0

    answered = len(
        answers
    )

    exam_start_timestamp = session.get(
        "exam_start_time"
    )

    now_datetime = datetime.now(
        ZoneInfo("Asia/Baku")
    )

    now_timestamp = now_datetime.timestamp()

    if exam_start_timestamp is not None:

        exam_start_timestamp = float(
            exam_start_timestamp
        )

        exam_start_datetime = datetime.fromtimestamp(
            exam_start_timestamp,
            ZoneInfo("Asia/Baku")
        )

        start_time_text = (
            exam_start_datetime.strftime(
                "%d.%m.%Y %H:%M:%S"
            )
        )

        elapsed_seconds = int(
            now_timestamp -
            exam_start_timestamp
        )

        if elapsed_seconds < 0:

            elapsed_seconds = 0

    else:

        start_time_text = "-"

        elapsed_seconds = 0

    end_time_text = (
        now_datetime.strftime(
            "%d.%m.%Y %H:%M:%S"
        )
    )

    duration_minutes = (
        elapsed_seconds // 60
    )

    duration_seconds = (
        elapsed_seconds % 60
    )

    duration_text = (
        f"{duration_minutes} dəq "
        f"{duration_seconds} san"
    )

    question_results = []

    for index, q in enumerate(
        current_questions
    ):

        selected = answers.get(
            str(index)
        )

        correct_letter = q["answer"]

        correct_text = q[
            correct_letter.lower()
        ]

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

            selected_text = (
                "Cavab verilməyib"
            )

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

    wrong = answered - correct

    unanswered = total - answered

    percent = round(
        correct / total * 100
    ) if total else 0

    name = session["name"]

    created_at = now_datetime.strftime(
        "%d.%m.%Y %H:%M:%S"
    )

    if answered == total:

        status = "Tamamlandı"

    else:

        status = (
            f"Yarımçıq bitirildi "
            f"({answered}/{total})"
        )

    try:

        sheet = get_sheet()

        all_values = sheet.get_all_values()

        number = len(
            all_values
        )

        row_data = [

            number,

            name,

            correct,

            total,

            wrong,

            f"{percent}%",

            created_at,

            status,

            duration_text,

            start_time_text,

            end_time_text

        ]

        sheet.append_row(
            row_data,
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

        duration_text=duration_text,

        duration_minutes=duration_minutes,

        duration_seconds=duration_seconds,

        start_time_text=start_time_text,

        end_time_text=end_time_text,

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

        questions=load_questions()

    )


# =========================================================
# IMPORT QUESTIONS
# =========================================================

@app.route(
    "/admin/import-questions",
    methods=["POST"]
)
@admin_required
def import_questions():

    file = request.files.get(
        "questions_file"
    )

    if not file:

        return redirect(
            url_for("admin")
        )

    filename = (
        file.filename or ""
    ).lower()

    new_questions = []

    try:

        # =====================================================
        # EXCEL
        # =====================================================

        if (
            filename.endswith(".xlsx")
            or filename.endswith(".xls")
        ):

            workbook = load_workbook(
                file,
                data_only=True
            )

            sheet = workbook.active

            for row in sheet.iter_rows(
                min_row=2,
                values_only=True
            ):

                if not row:
                    continue

                if row[0] is None:
                    continue

                question_text = str(
                    row[0]
                ).strip()

                if not question_text:
                    continue

                option_a = (
                    str(row[1]).strip()
                    if len(row) > 1
                    and row[1] is not None
                    else ""
                )

                option_b = (
                    str(row[2]).strip()
                    if len(row) > 2
                    and row[2] is not None
                    else ""
                )

                option_c = (
                    str(row[3]).strip()
                    if len(row) > 3
                    and row[3] is not None
                    else ""
                )

                option_d = (
                    str(row[4]).strip()
                    if len(row) > 4
                    and row[4] is not None
                    else ""
                )

                answer = (
                    str(row[5]).strip().upper()
                    if len(row) > 5
                    and row[5] is not None
                    else ""
                )

                if len(answer) > 1:

                    answer = answer[0]

                if answer not in [
                    "A",
                    "B",
                    "C",
                    "D"
                ]:

                    continue

                new_questions.append({

                    "question":
                        question_text,

                    "a":
                        option_a,

                    "b":
                        option_b,

                    "c":
                        option_c,

                    "d":
                        option_d,

                    "answer":
                        answer

                })


        # =====================================================
        # JSON
        # =====================================================

        elif filename.endswith(".json"):

            loaded_questions = json.load(
                file
            )

            if isinstance(
                loaded_questions,
                list
            ):

                for q in loaded_questions:

                    if not isinstance(
                        q,
                        dict
                    ):

                        continue

                    question_text = str(
                        q.get(
                            "question",
                            ""
                        )
                    ).strip()

                    if not question_text:

                        continue

                    answer = str(
                        q.get(
                            "answer",
                            ""
                        )
                    ).strip().upper()

                    if len(answer) > 1:

                        answer = answer[0]

                    if answer not in [
                        "A",
                        "B",
                        "C",
                        "D"
                    ]:

                        continue

                    new_questions.append({

                        "question":
                            question_text,

                        "a":
                            str(
                                q.get(
                                    "a",
                                    ""
                                )
                            ).strip(),

                        "b":
                            str(
                                q.get(
                                    "b",
                                    ""
                                )
                            ).strip(),

                        "c":
                            str(
                                q.get(
                                    "c",
                                    ""
                                )
                            ).strip(),

                        "d":
                            str(
                                q.get(
                                    "d",
                                    ""
                                )
                            ).strip(),

                        "answer":
                            answer

                    })


        else:

            print(
                "IMPORT ERROR: "
                "Fayl formatı dəstəklənmir."
            )

            return redirect(
                url_for("admin")
            )


        # =====================================================
        # BOŞ IMPORT OLARSA KÖHNƏ SUALLARI SİLMƏ
        # =====================================================

        if not new_questions:

            print(
                "IMPORT: Yeni sual tapılmadı. "
                "Mövcud suallar saxlanıldı."
            )

            return redirect(
                url_for("admin")
            )


        # =====================================================
        # GOOGLE SHEETS
        # =====================================================

        questions_sheet = get_questions_sheet()


        # ƏVVƏLKİ SUALLARI SİL
        questions_sheet.clear()


        # BAŞLIQLARI YENİDƏN YAZ
        questions_sheet.update(
            "A1:F1",
            [QUESTIONS_HEADERS]
        )


        # YENİ SUALLARI HAZIRLA
        rows = []

        for q in new_questions:

            rows.append([

                q["question"],

                q["a"],

                q["b"],

                q["c"],

                q["d"],

                q["answer"]

            ])


        # YENİ SUALLARI GOOGLE SHEETS-Ə YAZ
        questions_sheet.update(
            f"A2:F{len(rows) + 1}",
            rows
        )


        print(
            f"IMPORT UĞURLU: "
            f"{len(new_questions)} sual "
            f"Google Sheets-də saxlanıldı."
        )


    except Exception as e:

        print(
            "IMPORT ERROR:",
            str(e)
        )


    return redirect(
        url_for("admin")
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

    headers = [

        "№",

        "Ad və soyad",

        "Düzgün cavab",

        "Ümumi sual",

        "Səhv cavab",

        "Nəticə",

        "Tarix",

        "Status",

        "Müddət",

        "Başlama vaxtı",

        "Bitmə vaxtı"

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

        worksheet.cell(
            row=row_number,
            column=9,
            value=result.get(
                "Müddət",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=10,
            value=result.get(
                "Başlama vaxtı",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=11,
            value=result.get(
                "Bitmə vaxtı",
                ""
            )
        )

    widths = {

        "A": 8,

        "B": 30,

        "C": 18,

        "D": 18,

        "E": 15,

        "F": 15,

        "G": 25,

        "H": 30,

        "I": 18,

        "J": 25,

        "K": 25

    }

    for column, width in widths.items():

        worksheet.column_dimensions[
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
