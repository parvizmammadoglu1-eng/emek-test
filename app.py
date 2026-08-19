import os
import json
from datetime import datetime
from zoneinfo import ZoneInfo
from io import BytesIO
from functools import wraps

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
# BÖLMƏLƏR
# =========================================================

SECTIONS = [
    "I Bölmə",
    "II Bölmə",
    "III Bölmə",
    "IV Bölmə",
    "V Bölmə",
    "VI Bölmə",
    "VII Bölmə",
    "VIII Bölmə",
    "IX Bölmə",
    "X Bölmə",
    "XI Bölmə",
    "XII Bölmə",
    "XIII Bölmə"
]


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
            cols=12
        )

    headers = sheet.row_values(1)

    required_headers = [
        "№",
        "Ad və soyad",
        "Bölmə",
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
            "A1:L1",
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
    "Bölmə",
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
            rows=5000,
            cols=7
        )

        sheet.update(
            "A1:G1",
            [QUESTIONS_HEADERS]
        )

    return sheet


# =========================================================
# BÜTÜN SUALLARI YÜKLƏ
# =========================================================

def load_all_questions():

    try:

        sheet = get_questions_sheet()

        values = sheet.get_all_values()

        if len(values) <= 1:
            return []

        questions = []

        for row in values[1:]:

            if not row:
                continue

            if len(row) < 7:
                continue

            section = str(
                row[0]
            ).strip()

            question_text = str(
                row[1]
            ).strip()

            if not section or not question_text:
                continue

            answer = str(
                row[6]
            ).strip().upper()

            if answer not in [
                "A",
                "B",
                "C",
                "D"
            ]:
                continue

            questions.append({

                "section":
                    section,

                "question":
                    question_text,

                "a":
                    str(row[2]).strip(),

                "b":
                    str(row[3]).strip(),

                "c":
                    str(row[4]).strip(),

                "d":
                    str(row[5]).strip(),

                "answer":
                    answer
            })

        return questions

    except Exception as e:

        print(
            "LOAD ALL QUESTIONS ERROR:",
            str(e)
        )

        return []


# =========================================================
# SEÇİLMİŞ BÖLMƏNİN SUALLARINI YÜKLƏ
# =========================================================

def load_questions(section=None):

    questions = load_all_questions()

    if section is None:
        return questions

    section = str(
        section
    ).strip()

    return [
        q
        for q in questions
        if q.get(
            "section",
            ""
        ).strip() == section
    ]


# =========================================================
# KODLAR SHEET
# =========================================================

CODES_SHEET_NAME = "Kodlar"

CODES_HEADERS = [
    "Kod",
    "Bölmə",
    "Status",
    "Ad və soyad",
    "İstifadə vaxtı"
]


def get_codes_sheet():

    client = get_google_client()

    spreadsheet = client.open(
        GOOGLE_SHEET_NAME
    )

    try:

        sheet = spreadsheet.worksheet(
            CODES_SHEET_NAME
        )

    except gspread.WorksheetNotFound:

        sheet = spreadsheet.add_worksheet(
            title=CODES_SHEET_NAME,
            rows=2000,
            cols=5
        )

        sheet.update(
            "A1:E1",
            [CODES_HEADERS]
        )

    else:

        headers = sheet.row_values(1)

        if not headers:

            sheet.update(
                "A1:E1",
                [CODES_HEADERS]
            )

        else:

            changed = False

            for header in CODES_HEADERS:

                if header not in headers:

                    headers.append(
                        header
                    )

                    changed = True

            if changed:

                sheet.resize(
                    rows=max(
                        sheet.row_count,
                        2000
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
# KODU YOXLA VƏ BİR DƏFƏLİK İSTİFADƏ ET
# =========================================================

def use_access_code(
    code,
    selected_section,
    name
):

    code = str(
        code or ""
    ).strip()

    selected_section = str(
        selected_section or ""
    ).strip()

    name = str(
        name or ""
    ).strip()

    if not code:

        return False, (
            "Giriş kodunu daxil edin."
        )

    if selected_section not in SECTIONS:

        return False, (
            "Zəhmət olmasa bölmə seçin."
        )

    try:

        sheet = get_codes_sheet()

        values = sheet.get_all_values()

        if len(values) <= 1:

            return False, (
                "Hazırda aktiv giriş kodu yoxdur."
            )

        for row_number, row in enumerate(
            values[1:],
            start=2
        ):

            if not row:
                continue

            sheet_code = str(
                row[0]
                if len(row) > 0
                else ""
            ).strip()

            if sheet_code.upper() != code.upper():
                continue

            sheet_section = str(
                row[1]
                if len(row) > 1
                else ""
            ).strip()

            status = str(
                row[2]
                if len(row) > 2
                else ""
            ).strip().lower()

            # ---------------------------------------------
            # KOD ARTİQ İSTİFADƏ OLUNUB
            # ---------------------------------------------

            if status != "aktiv":

                return False, (
                    "Bu giriş kodu artıq istifadə olunub."
                )

            # ---------------------------------------------
            # BÖLMƏ YOXLAMASI
            # ---------------------------------------------

            if sheet_section != selected_section:

                return False, (
                    "Bu giriş kodu seçdiyiniz bölmə üçün "
                    "nəzərdə tutulmayıb."
                )

            # ---------------------------------------------
            # KOD DÜZGÜNDÜR
            # ---------------------------------------------

            current_time = datetime.now(
                ZoneInfo("Asia/Baku")
            ).strftime(
                "%d.%m.%Y %H:%M:%S"
            )

            sheet.update_cell(
                row_number,
                3,
                "İstifadə olunub"
            )

            sheet.update_cell(
                row_number,
                4,
                name
            )

            sheet.update_cell(
                row_number,
                5,
                current_time
            )

            print(
                f"GİRİŞ KODU İSTİFADƏ OLUNDU: "
                f"{code} | {selected_section}"
            )

            return True, ""

        return False, (
            "Giriş kodu yanlışdır."
        )

    except Exception as e:

        print(
            "ACCESS CODE ERROR:",
            str(e)
        )

        return False, (
            "Giriş kodu yoxlanılarkən xəta baş verdi. "
            "Bir qədər sonra yenidən cəhd edin."
        )


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

        return function(
            *args,
            **kwargs
        )

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

        access_code = request.form.get(
            "access_code",
            ""
        ).strip()

        selected_section = request.form.get(
            "section",
            ""
        ).strip()

        if not name:

            return render_template(
                "home.html",
                error="Ad və soyad daxil edin.",
                name=name,
                access_code=access_code,
                selected_section=selected_section,
                sections=SECTIONS
            )

        if selected_section not in SECTIONS:

            return render_template(
                "home.html",
                error="Bölmə seçin.",
                name=name,
                access_code=access_code,
                selected_section=selected_section,
                sections=SECTIONS
            )

        if not access_code:

            return render_template(
                "home.html",
                error="Giriş kodunu daxil edin.",
                name=name,
                access_code=access_code,
                selected_section=selected_section,
                sections=SECTIONS
            )

        section_questions = load_questions(
            selected_section
        )

        if not section_questions:

            return render_template(
                "home.html",
                error=(
                    f"{selected_section} üzrə hazırda "
                    f"sual mövcud deyil."
                ),
                name=name,
                access_code=access_code,
                selected_section=selected_section,
                sections=SECTIONS
            )

        code_valid, error_message = use_access_code(
            access_code,
            selected_section,
            name
        )

        if not code_valid:

            return render_template(
                "home.html",
                error=error_message,
                name=name,
                access_code=access_code,
                selected_section=selected_section,
                sections=SECTIONS
            )

        session.clear()

        session["name"] = name

        session["access_code"] = access_code

        session["section"] = selected_section

        session["answers"] = {}

        session["exam_finished"] = False

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
        "home.html",
        sections=SECTIONS,
        selected_section=""
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

    if "access_code" not in session:
        return redirect(
            url_for("home")
        )

    if "section" not in session:
        return redirect(
            url_for("home")
        )

    if session.get("exam_finished"):

        return redirect(
            url_for("finish")
        )

    selected_section = session.get(
        "section"
    )

    current_questions = load_questions(
        selected_section
    )

    total = len(
        current_questions
    )

    if total == 0:

        return render_template(
            "home.html",
            error=(
                f"{selected_section} üzrə "
                f"hazırda sual yoxdur."
            ),
            sections=SECTIONS
        )

    if n < 0:

        return redirect(
            url_for(
                "question",
                n=0
            )
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

        section=selected_section,

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

    if "access_code" not in session:
        return redirect(
            url_for("home")
        )

    if "section" not in session:
        return redirect(
            url_for("home")
        )

    if session.get("exam_finished"):

        return render_template(
            "finish.html",
            name=session.get("name"),
            section=session.get("section"),
            correct=session.get("finish_correct", 0),
            total=session.get("finish_total", 0),
            percent=session.get("finish_percent", 0),
            answered=session.get("finish_answered", 0),
            wrong=session.get("finish_wrong", 0),
            unanswered=session.get("finish_unanswered", 0),
            status=session.get("finish_status", ""),
            duration_text=session.get(
                "finish_duration_text",
                ""
            ),
            duration_minutes=session.get(
                "finish_duration_minutes",
                0
            ),
            duration_seconds=session.get(
                "finish_duration_seconds",
                0
            ),
            start_time_text=session.get(
                "finish_start_time_text",
                "-"
            ),
            end_time_text=session.get(
                "finish_end_time_text",
                "-"
            ),
            question_results=session.get(
                "finish_question_results",
                []
            )
        )

    answers = session.get(
        "answers",
        {}
    )

    selected_section = session.get(
        "section"
    )

    current_questions = load_questions(
        selected_section
    )

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

    # =====================================================
    # GOOGLE SHEETS NƏTİCƏ
    # =====================================================

    try:

        sheet = get_sheet()

        all_values = sheet.get_all_values()

        number = len(
            all_values
        )

        row_data = [

            number,

            name,

            selected_section,

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

    # =====================================================
    # SESSION
    # =====================================================

    session["exam_finished"] = True

    session["finish_correct"] = correct
    session["finish_total"] = total
    session["finish_percent"] = percent
    session["finish_answered"] = answered
    session["finish_wrong"] = wrong
    session["finish_unanswered"] = unanswered
    session["finish_status"] = status
    session["finish_duration_text"] = duration_text
    session["finish_duration_minutes"] = duration_minutes
    session["finish_duration_seconds"] = duration_seconds
    session["finish_start_time_text"] = start_time_text
    session["finish_end_time_text"] = end_time_text
    session["finish_question_results"] = question_results

    session.modified = True

    return render_template(

        "finish.html",

        name=name,

        section=selected_section,

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

    # -----------------------------------------------------
    # NƏTİCƏLƏR
    # -----------------------------------------------------

    try:

        sheet = get_sheet()

        values = sheet.get_all_records()

    except Exception as e:

        print(
            "ADMIN GOOGLE SHEETS ERROR:",
            str(e)
        )

        values = []

    # -----------------------------------------------------
    # SUALLAR
    # -----------------------------------------------------

    questions = load_all_questions()

    # -----------------------------------------------------
    # KODLAR
    # -----------------------------------------------------

    try:

        codes_sheet = get_codes_sheet()

        codes = codes_sheet.get_all_records()

    except Exception as e:

        print(
            "ADMIN CODES ERROR:",
            str(e)
        )

        codes = []

    # -----------------------------------------------------
    # PANEL
    # -----------------------------------------------------

    return render_template(

        "admin.html",

        results=values,

        questions=questions,

        sections=SECTIONS,

        codes=codes

    )


# =========================================================
# ADMIN - YENİ GİRİŞ KODU ƏLAVƏ ET
# =========================================================

@app.route(
    "/admin/add-code",
    methods=["POST"]
)
@admin_required
def admin_add_code():

    code = request.form.get(
        "code",
        ""
    ).strip()

    section = request.form.get(
        "section",
        ""
    ).strip()

    if not code:

        return redirect(
            url_for("admin")
        )

    if section not in SECTIONS:

        return redirect(
            url_for("admin")
        )

    try:

        sheet = get_codes_sheet()

        values = sheet.get_all_values()

        # -------------------------------------------------
        # EYNİ KODU YENİDƏN ƏLAVƏ ETMƏ
        # -------------------------------------------------

        for row in values[1:]:

            if not row:
                continue

            existing_code = str(
                row[0]
                if len(row) > 0
                else ""
            ).strip()

            if existing_code.upper() == code.upper():

                print(
                    "ADD CODE: Bu kod artıq mövcuddur:",
                    code
                )

                return redirect(
                    url_for("admin")
                )

        # -------------------------------------------------
        # YENİ KOD
        # -------------------------------------------------

        sheet.append_row(
            [
                code,
                section,
                "Aktiv",
                "",
                ""
            ],
            value_input_option="USER_ENTERED"
        )

        print(
            f"YENİ KOD ƏLAVƏ EDİLDİ: "
            f"{code} | {section}"
        )

    except Exception as e:

        print(
            "ADD CODE ERROR:",
            str(e)
        )

    return redirect(
        url_for("admin")
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

        # =================================================
        # EXCEL
        # =================================================

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

                section_value = (
                    str(row[0]).strip()
                    if len(row) > 0
                    and row[0] is not None
                    else ""
                )

                section = normalize_section(
                    section_value
                )

                if section not in SECTIONS:

                    print(
                        "IMPORT: Yanlış bölmə:",
                        section_value
                    )

                    continue

                question_text = (
                    str(row[1]).strip()
                    if len(row) > 1
                    and row[1] is not None
                    else ""
                )

                if not question_text:
                    continue

                option_a = (
                    str(row[2]).strip()
                    if len(row) > 2
                    and row[2] is not None
                    else ""
                )

                option_b = (
                    str(row[3]).strip()
                    if len(row) > 3
                    and row[3] is not None
                    else ""
                )

                option_c = (
                    str(row[4]).strip()
                    if len(row) > 4
                    and row[4] is not None
                    else ""
                )

                option_d = (
                    str(row[5]).strip()
                    if len(row) > 5
                    and row[5] is not None
                    else ""
                )

                answer = (
                    str(row[6]).strip().upper()
                    if len(row) > 6
                    and row[6] is not None
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

                    "section":
                        section,

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

        # =================================================
        # JSON
        # =================================================

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

                    section = normalize_section(
                        str(
                            q.get(
                                "section",
                                ""
                            )
                        ).strip()
                    )

                    if section not in SECTIONS:
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

                        "section":
                            section,

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

        # =================================================
        # BOŞ IMPORT
        # =================================================

        if not new_questions:

            print(
                "IMPORT: Yeni sual tapılmadı. "
                "Mövcud suallar saxlanıldı."
            )

            return redirect(
                url_for("admin")
            )

        # =================================================
        # SUALLAR SHEET
        # =================================================

        questions_sheet = get_questions_sheet()

        # Yalnız SUALLAR sheet-i silinir.
        # KODLAR sheet-i toxunulmaz qalır.

        questions_sheet.clear()

        questions_sheet.update(
            "A1:G1",
            [QUESTIONS_HEADERS]
        )

        rows = []

        for q in new_questions:

            rows.append([

                q["section"],

                q["question"],

                q["a"],

                q["b"],

                q["c"],

                q["d"],

                q["answer"]

            ])

        questions_sheet.update(
            f"A2:G{len(rows) + 1}",
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
# BÖLMƏ NORMALİZASİYASI
# =========================================================

def normalize_section(value):

    value = str(
        value or ""
    ).strip()

    if not value:
        return ""

    numeric_sections = {

        "1": "I Bölmə",
        "2": "II Bölmə",
        "3": "III Bölmə",
        "4": "IV Bölmə",
        "5": "V Bölmə",
        "6": "VI Bölmə",
        "7": "VII Bölmə",
        "8": "VIII Bölmə",
        "9": "IX Bölmə",
        "10": "X Bölmə",
        "11": "XI Bölmə",
        "12": "XII Bölmə",
        "13": "XIII Bölmə"

    }

    if value in numeric_sections:

        return numeric_sections[value]

    cleaned = (
        value
        .lower()
        .replace("bölmə", "")
        .replace("bölm", "")
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

    if cleaned in numeric_sections:

        return numeric_sections[cleaned]

    for section in SECTIONS:

        if value.lower() == section.lower():

            return section

    return value


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
        "Bölmə",
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
                "Bölmə",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=4,
            value=result.get(
                "Düzgün cavab",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=5,
            value=result.get(
                "Ümumi sual",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=6,
            value=result.get(
                "Səhv cavab",
                0
            )
        )

        worksheet.cell(
            row=row_number,
            column=7,
            value=result.get(
                "Nəticə",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=8,
            value=result.get(
                "Tarix",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=9,
            value=result.get(
                "Status",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=10,
            value=result.get(
                "Müddət",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=11,
            value=result.get(
                "Başlama vaxtı",
                ""
            )
        )

        worksheet.cell(
            row=row_number,
            column=12,
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
        "E": 18,
        "F": 15,
        "G": 15,
        "H": 25,
        "I": 30,
        "J": 18,
        "K": 25,
        "L": 25

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
        "emek_mecellesi_test_neticeleri.xlsx"
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
