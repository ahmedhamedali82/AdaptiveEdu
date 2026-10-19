#!/usr/bin/env python3
"""AdaptiveEdu documentation + video script (live bot: Gemini + Excel + 3 Sheets)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path("/Users/ahmedhamed/Documents/AI Agent/Technical/Graduation_Project/AdaptiveEdu")
DEMO = ROOT / "diagrams" / "demo"
OUT_DOC = ROOT / "AdaptiveEdu_Project_Documentation.docx"
OUT_SCRIPT = ROOT / "AdaptiveEdu_Video_demo_steps.docx"

NAVY = RGBColor(15, 39, 68)
TEAL = RGBColor(14, 140, 126)
MUTED = RGBColor(90, 102, 118)
DARK = RGBColor(28, 36, 48)
GREEN = RGBColor(46, 140, 90)
RED = RGBColor(185, 28, 28)

GITHUB = "https://github.com/ahmedhamedali82/AdaptiveEdu"
BOT = "https://t.me/AdaptiveEdubot"
SHEET_STU = "https://docs.google.com/spreadsheets/d/1ddIwenYjiwKryP7Pogxe-7qsjKnTmtfCtAlDoplSn5Q"
SHEET_LES = "https://docs.google.com/spreadsheets/d/1akeM65odle57kxtR7M5LqFMnxCsHX3M1cJKKltExFEg"
SHEET_PRG = "https://docs.google.com/spreadsheets/d/1II6ckyiKFucKsC38PC_0a9tH1h8H0JZe3qaclLxCExY"

SUBJECTS = [
    ("1", "Arabic Language"),
    ("2", "Mathematics"),
    ("3", "Integrated Sciences"),
    ("4", "Philosophy & Logic"),
    ("5", "First Foreign Language / English"),
    ("6", "Egyptian History"),
    ("7", "Programming & Artificial Intelligence"),
    ("8", "Physical Education"),
]


def set_run(run, size=12, bold=False, color=DARK, name="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:cs"), "Arial")


def shade_cell(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def add_p(doc, text, size=12, bold=False, color=DARK, space=8, align="left", rtl=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right" or rtl:
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if rtl:
        pPr = p._p.get_or_add_pPr()
        bidi = OxmlElement("w:bidi")
        bidi.set(qn("w:val"), "1")
        pPr.append(bidi)
    run = p.add_run(text)
    set_run(run, size=size, bold=bold, color=color, name="Arial" if rtl else "Calibri")
    return p


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16 if level == 1 else 10)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    set_run(run, size=18 if level == 1 else 14, bold=True, color=NAVY if level == 1 else TEAL)
    return p


def heading_ar(doc, text):
    add_p(doc, text, size=16, bold=True, color=NAVY, space=8, rtl=True)


def bullets(doc, items, rtl=False):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        if rtl:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pPr = p._p.get_or_add_pPr()
            bidi = OxmlElement("w:bidi")
            bidi.set(qn("w:val"), "1")
            pPr.append(bidi)
        run = p.add_run(item)
        set_run(run, size=12, name="Arial" if rtl else "Calibri")


def table(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        set_run(run, size=11, bold=True, color=RGBColor(255, 255, 255))
        shade_cell(cell, "0F2744")
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = t.rows[r_i + 1].cells[c_i]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            set_run(run, size=11)
            if r_i % 2 == 0:
                shade_cell(cell, "F6F8FB")
    doc.add_paragraph()


def setup(doc):
    for sec in doc.sections:
        sec.top_margin = Cm(1.8)
        sec.bottom_margin = Cm(1.8)
        sec.left_margin = Cm(1.8)
        sec.right_margin = Cm(1.8)


def add_caption(doc, text):
    add_p(doc, text, size=10, color=MUTED, space=12, align="center")


def add_img(doc, path, width_in=5.8, caption=None):
    p = Path(path)
    if not p.exists():
        add_p(doc, f"[Missing image: {p.name}]", size=11, color=RED, space=6)
        return
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_after = Pt(4)
    run = para.add_run()
    run.add_picture(str(p), width=Inches(width_in))
    if caption:
        add_caption(doc, caption)


def quote_box(doc, title, body, rtl=False):
    add_p(doc, title, size=12, bold=True, color=TEAL, space=2)
    add_p(doc, body, size=11, color=DARK, space=10, rtl=rtl)


# ── Video script ─────────────────────────────────────────────────────────────


def build_script():
    """Short live-demo shot list only."""
    doc = Document()
    setup(doc)
    add_p(doc, "AdaptiveEdu — Video demo steps", size=22, bold=True, color=NAVY, space=4)
    add_p(doc, "2–5 min screen recording of the live bot  ·  " + BOT, size=12, color=MUTED, space=10)

    heading(doc, "Before record")
    bullets(
        doc,
        [
            "n8n Active · Telegram as ST005 · Gmail open · file ST005_4_1.xlsx ready (Pass answers).",
            "Layout: Telegram + Gmail on screen.",
        ],
    )

    heading(doc, "Demo steps (do this on camera)")
    steps = [
        ("1", "Open AdaptiveEdubot", "Show the bot name."),
        ("2", "Send ST005", "Wait for subject list 1–8."),
        ("3", "Send 4", "Wait for lesson / send notice (Philosophy)."),
        ("4", "Open lesson email", "Show PDF attachment + name ST005_4_1.xlsx (open PDF briefly)."),
        ("5", "Show Excel file", "One row, 10 capital letters, no header, correct file name."),
        ("6", "Upload Excel on Telegram", "Wait for grading message."),
        ("7", "Open result email", "Show score + Pass or Fail."),
        ("8", "Open student_progress", "Show the row (grade / succed), then stop."),
    ]
    table(doc, ["#", "Action", "Show on screen"], steps)

    heading(doc, "Say (short)")
    bullets(
        doc,
        [
            "“This is AdaptiveEdu — live Telegram study bot.”",
            "“ID ST005 → subject 4 → lesson by email.”",
            "“Answers file ST005_4_1.xlsx — one row, ten letters.”",
            "“Result email: Pass if more than six out of ten. Saved in the sheet. Thank you.”",
        ],
    )

    heading_ar(doc, "بالعربي — نفس الخطوات")
    bullets(
        doc,
        [
            "افتح البوت → ابعت ST005 → ابعت 4",
            "افتح إيميل الدرس (مرفق PDF + اسم ST005_4_1.xlsx)",
            "اعرض الإكسل → ارفعه على تيليجرام",
            "افتح إيميل النتيجة + شيت student_progress",
        ],
        rtl=True,
    )

    doc.save(OUT_SCRIPT)
    print(OUT_SCRIPT)



# ── Full documentation ───────────────────────────────────────────────────────


def build_docs():
    doc = Document()
    setup(doc)

    # Cover
    add_p(doc, "AdaptiveEdu", size=26, bold=True, color=NAVY, space=2)
    add_p(doc, "Project documentation — DEPI Graduation Project", size=16, color=TEAL, space=4)
    add_p(
        doc,
        "Live Telegram study bot for Egyptian 1st Secondary  ·  Gemini lessons  ·  Excel quiz  ·  Gmail Pass/Fail",
        size=12,
        color=MUTED,
    )
    add_p(doc, BOT, size=12, bold=True, color=TEAL, space=2)
    add_p(doc, GITHUB, size=12, bold=True, color=TEAL, space=12)

    heading(doc, "1. Abstract")
    add_p(
        doc,
        "AdaptiveEdu is a live Telegram bot (t.me/AdaptiveEdubot) that registers a student by ID, "
        "lets them choose a subject, generates the next lesson with Google Gemini as a PDF plus ten MCQ model answers, "
        "emails the PDF as an attachment with the required Excel answer-file name, then grades the uploaded Excel file. "
        "Pass means grade greater than 6 out of 10. Fail means 6 or less — the same lesson stays open until Pass. "
        "n8n workflow ADAPTIVE_14 (about 46 nodes) orchestrates Telegram, three Google Sheets, Drive, Gmail, Gemini, and PDFShift (HTML → PDF). "
        "The student never opens n8n. Telegram and email are the student screens.",
    )

    heading(doc, "2. Problem")
    bullets(
        doc,
        [
            "One classroom, one pace, one book.",
            "A student who fails a lesson still moves on with the class.",
            "A teacher cannot follow every student by hand after school.",
            "We need a simple rule: finish this lesson — with a scored quiz — before the next one.",
        ],
    )

    heading(doc, "3. Objectives")
    bullets(
        doc,
        [
            "Register each student once and bind their Telegram chat ID.",
            "Send only the next unanswered / unpassed lesson for the chosen subject.",
            "Generate lesson content with Gemini (Markdown → HTML → PDF) and ten model answers.",
            "Accept answers only as Excel: one row, ten capital letters, no header, exact file name.",
            "Pass if grade > 6 / 10; Fail otherwise; block the next lesson until Pass.",
            "Email lesson materials and Pass/Fail results; log everything in student_progress.",
            "Keep the stack clear: Telegram + n8n + 3 Sheets + Gemini + Gmail + Drive.",
        ],
    )

    heading(doc, "4. Users")
    table(
        doc,
        ["User", "What they do"],
        [
            ["Student", "Sends ST### ID, picks subject code 1–8, opens the emailed PDF attachment, uploads Excel answers."],
            ["Teacher / team", "Checks Students_Data, Lessons_Data, and student_progress. Helps stuck students by hand."],
            ["System (n8n)", "Registers, generates lesson, emails, grades Excel, updates progress, replies Pass/Fail."],
        ],
    )

    heading(doc, "5. How the student uses it (8 steps)")
    bullets(
        doc,
        [
            "Open https://t.me/AdaptiveEdubot",
            "Send student ID, example: ST005",
            "Bot replies with subject list 1–8 (and tip to upload an answers file).",
            "Send subject code, example: 4 = Philosophy & Logic",
            "Receive Telegram notice + Gmail with PDF attachment and required Excel name (example: ST005_4_1.xlsx)",
            "Study the PDF, then fill one Excel row: 10 capital letters A–D (or as required), no title row",
            "Upload the Excel file on Telegram with the exact name",
            "Receive grading confirmation on Telegram + result email: Pass or Fail",
        ],
    )

    heading(doc, "6. Architecture (three paths)")
    add_p(
        doc,
        "Telegram Trigger receives every message. The workflow branches on message type:",
        space=6,
    )
    table(
        doc,
        ["Path", "Input", "Main outcome"],
        [
            ["A — Register", "Student ID text (ST###)", "Bind telegram_chat_id, show subjects, set waiting flag"],
            ["B — Lesson", "Subject code 1–8", "Pick next lesson → Gemini → PDF → email + Telegram → append progress"],
            ["C — Grade", "Excel document", "Validate name → parse row → compare ANS_1…ANS_10 → Pass/Fail email"],
        ],
    )
    add_p(
        doc,
        "If a previous lesson was delivered but not answered (or not passed), Path B blocks a new lesson and asks the student to finish the pending quiz first.",
    )
    add_img(
        doc,
        DEMO / "workflow_n8n.jpeg",
        width_in=6.2,
        caption="Figure 1 — Live n8n canvas: ADAPTIVE_14 (~46 nodes).",
    )

    heading(doc, "7. Tools and what each one does")
    table(
        doc,
        ["Tool", "Role"],
        [
            ["Telegram", "Only student chat UI. Bot: AdaptiveEdubot."],
            ["n8n", "Orchestration. Import AdaptiveEdu.json (workflow name ADAPTIVE_14)."],
            ["Google Sheets ×3", "Students_Data, Lessons_Data, student_progress."],
            ["Google Gemini (AI Agent)", "Writes Markdown lesson + 10 model answers."],
            ["PDFShift (HTTP Request)", "Converts formatted lesson HTML to a binary PDF file."],
            ["Google Drive", "Stores / moves answer files for extract-from-file."],
            ["Gmail", "Sends lesson email with PDF attachment + Excel name, and Pass/Fail result emails."],
            ["GitHub", GITHUB],
            ["PowerPoint", "AdaptiveEdu_Project_Plan.pptx — journey, n8n steps, live screenshots."],
        ],
    )

    heading(doc, "8. Data — Students_Data")
    add_p(doc, SHEET_STU, size=11, color=TEAL)
    add_p(doc, "Tab name: Students_Data. n8n matches on student_id and telegram_chat_id.")
    table(
        doc,
        ["Column", "Who fills it", "Meaning"],
        [
            ["student_id", "Team (before demo)", "Example ST005. What the student types."],
            ["student_name", "Team", "Used in emails (Dear …)."],
            ["phone", "Team (optional)", "Contact."],
            ["telegram_chat_id", "n8n", "Saved on first successful register."],
            ["email", "Team", "Where lesson and result emails are sent."],
            ["otp", "Optional", "Not required for the main demo path."],
            ["status", "n8n", "Waiting flag (0/1): next message treated as subject/file vs register."],
        ],
    )

    heading(doc, "9. Data — Lessons_Data")
    add_p(doc, SHEET_LES, size=11, color=TEAL)
    add_p(doc, "Two tabs. Do not rename them.")
    table(
        doc,
        ["Tab", "n8n action", "Columns / content"],
        [
            ["subject", "Read", "SUBJECT_CODE, subject — codes 1 to 8"],
            ["Lessons_Data", "Read", "subject_code, subject, lesson_code, lesson_name, optional file_name"],
        ],
    )
    table(doc, ["Code", "Subject"], SUBJECTS)
    add_p(
        doc,
        "Demo lesson used in screenshots: subject 4 Philosophy & Logic → lesson_code 1 Thinking Principles.",
    )

    heading(doc, "10. Data — student_progress")
    add_p(doc, SHEET_PRG, size=11, color=TEAL)
    add_p(
        doc,
        "History of delivered lessons and graded attempts. "
        "Column headers used by the live workflow: sudent_id and succed.",
    )
    table(
        doc,
        ["Column", "Meaning"],
        [
            ["sudent_id", "Student ID"],
            ["subject_code", "1–8"],
            ["lesson_id / lesson_code", "Lesson number for that subject"],
            ["date", "Attempt / delivery date"],
            ["grade", "Number correct out of 10"],
            ["succed", "1 = Pass, 0 = Fail"],
            ["FILE_NAME", "Expected Excel name, e.g. ST005_4_1.xlsx"],
            ["ANS_1 … ANS_10", "Model answers stored when the lesson is generated"],
        ],
    )

    heading(doc, "11. Quiz Excel file — exact style")
    bullets(
        doc,
        [
            "File name pattern: ST{student_id}_{subject_code}_{lesson_code}.xlsx",
            "Example: ST005_4_1.xlsx = student ST005, Philosophy (4), lesson 1",
            "One record (one row), ten columns, capital letters, no title/header row",
            "Wrong name → bot refuses the file (file_error path)",
        ],
    )
    add_img(
        doc,
        DEMO / "quiz_excel_style.png",
        width_in=6.0,
        caption="Figure 2 — Sample answer sheet style (ST005_4_1.xlsx).",
    )

    heading(doc, "12. n8n workflow — step detail")
    add_p(doc, "Workflow name: ADAPTIVE_14. Import file: AdaptiveEdu.json.", space=6)

    heading(doc, "12.A  Register path", level=2)
    table(
        doc,
        ["Step", "Node / action", "Job"],
        [
            ["1", "Telegram Trigger", "New message arrives"],
            ["2", "Get row_by_chatid", "Look up Students_Data by chat ID"],
            ["3", "CHECH_WAITNG", "If not waiting → register / menu path"],
            ["4", "Search / check_student_id", "Match text like ST005"],
            ["5", "Update_chat_id", "Save Telegram id once"],
            ["6", "Get Subjects + subject_display", "Show list 1–8 + upload tip"],
            ["7", "Update_waitig_1", "status = waiting for subject or file"],
        ],
    )

    heading(doc, "12.B  Lesson path", level=2)
    table(
        doc,
        ["Step", "Node / action", "Job"],
        [
            ["1", "subject_check", "Text equals a SUBJECT_CODE"],
            ["2", "GET_PROGRESS + Get Lessons Data", "Find last passed / next lesson"],
            ["3", "get_file (Code)", "Pick next lesson_code / lesson_name"],
            ["4", "check_lesson_delivry / If_delivery", "Block if unanswered pending lesson"],
            ["5", "AI Agent + Gemini", "Markdown lesson + 10 answers"],
            ["6", "Format Lesson HTML + HTTP Request (PDFShift)", "Build binary PDF (property: pdf)"],
            ["7", "Send a message (Gmail)", "Email PDF attachment + required Excel name"],
            ["8", "add_lesson + send_file", "Append progress row + Telegram notice"],
        ],
    )

    heading(doc, "12.C  Grade path", level=2)
    table(
        doc,
        ["Step", "Node / action", "Job"],
        [
            ["1", "If (document)", "Message has a file"],
            ["2", "file_check", "Name must match ST###_sub_les.xlsx"],
            ["3", "Drive + Extract from File", "Upload / download / parse one row"],
            ["4", "Code compare", "Student letters vs ANS_1…ANS_10"],
            ["5", "Score", "grade = correct count / 10"],
            ["6", "Pass rule", "Pass if grade > 6 else Fail"],
            ["7", "Update row in sheet", "Write grade + succed flag"],
            ["8", "Gmail + Telegram", "Result email + restudy tip if Fail"],
        ],
    )

    heading(doc, "13. Pass / Fail rule")
    bullets(
        doc,
        [
            "Pass: grade > 6 out of 10. student_progress.succed = 1. Email says “you are Pass”. Next subject request opens the next lesson_code.",
            "Fail: grade ≤ 6. succed = 0. Email says “you are Fail”. Telegram asks to restudy and resend answers. Same lesson stays open.",
            "Unknown student ID → Arabic error: student code incorrect / not registered.",
            "Unknown subject code → Arabic error: subject code does not exist.",
            "Wrong Excel name → file refused.",
            "Pending unanswered lesson → new subject request blocked until the open quiz is submitted.",
        ],
    )

    heading(doc, "14. Live evidence — Telegram bot")
    add_p(
        doc,
        "After a valid ID, the bot lists subjects and offers to send an answers file. "
        "Wrong codes prove the Sheets validation is live.",
    )
    add_img(
        doc,
        DEMO / "bot_chat.jpeg",
        width_in=3.4,
        caption="Figure 3 — Bot subject list after registration (partial list visible) + upload tip in Arabic.",
    )
    add_img(
        doc,
        DEMO / "bot_errors.jpg",
        width_in=3.4,
        caption="Figure 4 — Validation errors: wrong subject code, then wrong/unknown student ID (duplicate replies possible on retry).",
    )
    quote_box(
        doc,
        "Arabic — wrong subject",
        "كود الماده الدراسيه غير موجود من فضلك ارسل كود صحيح",
        rtl=True,
    )
    quote_box(
        doc,
        "Arabic — wrong / unknown student ID",
        "عذراً، كود الطالب الذي أدخلته غير صحيح أو غير مسجل لدينا. يرجى التأكد من الرقم وإعادة المحاولة.",
        rtl=True,
    )
    add_img(doc, DEMO / "msg_welcome.png", width_in=5.5, caption="Figure 5 — Welcome / subject-list message card.")
    add_img(doc, DEMO / "msg_lesson_sent.png", width_in=5.5, caption="Figure 6 — Telegram notice after a lesson is generated.")
    add_img(doc, DEMO / "msg_graded.png", width_in=5.5, caption="Figure 7 — Telegram confirmation after Excel upload (before / with result email).")

    heading(doc, "15. Live evidence — lesson email")
    add_p(doc, "Subject line example: Thinking Principles", size=11, color=MUTED)
    quote_box(
        doc,
        "Email body (lesson delivery)",
        "Dear Ahmed Hamed,\n\n"
        "Kindly find the attached file from AdaptiveEdu system , regarding to\n"
        "the subject : Philosophy & Logic\n"
        "the lesson name is :Thinking Principles\n"
        "please download this file first then\n"
        "study it well and answer the attached quiz then send the answer in\n"
        "excel file with the name ST005_4_1.xlsx to correct it otherwise it\n"
        "will be refused\n"
        "note that : the answer is with capital letter in one record and 10\n"
        "columns in the excel sheet without title",
    )
    add_img(
        doc,
        DEMO / "email_lesson.png",
        width_in=5.8,
        caption="Figure 8 — Lesson delivery email (Philosophy & Logic → Thinking Principles → ST005_4_1.xlsx).",
    )

    heading(doc, "16. Live evidence — Fail result email")
    add_p(doc, "Subject line: ST005_4_1.xlsx_result", size=11, color=MUTED)
    quote_box(
        doc,
        "Email body (Fail — score 2 / 10)",
        "Dear Ahmed Hamed,\n\n"
        "your quiz is corected and your result is 2 from 10\n\n"
        "and you are Fail",
    )
    add_p(
        doc,
        "What happens next: student_progress.succed = 0; same lesson stays open; student must restudy and upload again.",
        size=11,
        color=RED,
    )
    add_img(
        doc,
        DEMO / "email_fail.png",
        width_in=5.5,
        caption="Figure 9 — Fail notification email (2 from 10).",
    )

    heading(doc, "17. Live evidence — Pass result email")
    add_p(doc, "Subject line: ST005_4_1.xlsx_result", size=11, color=MUTED)
    quote_box(
        doc,
        "Email body (Pass — score 9 / 10)",
        "Dear Ahmed Hamed,\n\n"
        "your quiz is corected and your result is 9 from 10\n\n"
        "and you are Pass",
    )
    add_p(
        doc,
        "What happens next: student_progress.succed = 1; next subject request can open lesson_code + 1.",
        size=11,
        color=GREEN,
    )
    add_img(
        doc,
        DEMO / "email_pass.png",
        width_in=5.5,
        caption="Figure 10 — Pass notification email (9 from 10).",
    )
    heading(doc, "18. How to run (setup)")
    bullets(
        doc,
        [
            "Create / use Telegram bot AdaptiveEdubot (BotFather token stays in n8n credentials — never in GitHub).",
            "In n8n: Import AdaptiveEdu.json (ADAPTIVE_14).",
            "Attach Telegram credential to Telegram Trigger and all Telegram send nodes.",
            "Attach Google Sheets credential to every Sheets node; confirm the three sheet IDs and tab names.",
            "Attach Gmail, Google Drive, Gemini, and set the PDFShift API key on the HTTP Request node.",
            "Activate the workflow. Send any message once so the Telegram Trigger registers.",
            "Demo student must exist in Students_Data (example ST005 / Ahmed Hamed) with a valid email.",
        ],
    )

    heading(doc, "19. Demo script (~3 minutes live)")
    table(
        doc,
        ["Step", "Do this", "What the panel should see"],
        [
            ["1", "Telegram: ST005", "Subject list 1–8"],
            ["2", "Telegram: 4", "Lesson generated notice"],
            ["3", "Open Gmail", "PDF attachment + required name ST005_4_1.xlsx"],
            ["4", "Show Excel", "One row, 10 capital letters, no header"],
            ["5", "Upload Excel on Telegram", "Grading confirmation"],
            ["6", "Show Pass or Fail email", "Then open student_progress row"],
        ],
    )

    heading(doc, "20. Expected result")
    bullets(
        doc,
        [
            "A student can complete one full loop: ID → subject → lesson email with PDF attachment → Excel → Pass/Fail email.",
            "Pass and Fail both work and are visible in email + student_progress.",
            "Wrong ID / subject / file name produce clear rejection messages.",
            "The panel can open the bot, the PPT, the JSON, and the three live sheets.",
        ],
    )

    heading(doc, "21. What we built vs what we did not build")
    table(
        doc,
        ["In scope (this delivery)", "Out of scope"],
        [
            ["Telegram bot + n8n ADAPTIVE_14 + 3 Sheets", "Mobile app, LMS, or school login portal"],
            ["Gemini-generated lesson PDF + 10 MCQs", "Full national exam bank authored by hand"],
            ["Excel answer upload + Pass if grade > 6", "Voice answers or in-chat letter quizzes as the main path"],
            ["Gmail lesson PDF attachment + Pass/Fail emails", "Automatic teacher dashboard UI"],
            ["Pending-lesson lock until Pass", "Every subject fully authored as long textbooks"],
        ],
    )

    heading(doc, "22. Deliverables checklist")
    bullets(
        doc,
        [
            "Working Telegram bot: https://t.me/AdaptiveEdubot",
            "n8n JSON: AdaptiveEdu.json (ADAPTIVE_14)",
            "Three sheets: Students_Data · Lessons_Data · student_progress",
            "Presentation: AdaptiveEdu_Project_Plan.pptx",
            "Documentation: AdaptiveEdu_Project_Documentation.docx",
            "Video demo steps: AdaptiveEdu_Video_demo_steps.docx",
            "GitHub: " + GITHUB,
            "Video 2–5 minutes following the demo steps",
        ],
    )

    heading(doc, "23. Project files")
    table(
        doc,
        ["File", "Purpose"],
        [
            ["README.md", "Short landing page"],
            ["AdaptiveEdu.json", "n8n workflow to import (ADAPTIVE_14)"],
            ["diagrams/demo/", "Bot, email, Excel, and canvas screenshots used in PPT + docs"],
            ["AdaptiveEdu_Project_Plan.pptx", "Panel slides"],
            ["AdaptiveEdu_Video_demo_steps.docx", "Video demo shot list"],
            ["AdaptiveEdu_Project_Documentation.docx", "This full document"],
        ],
    )

    heading(doc, "24. Risks and simple mitigations")
    bullets(
        doc,
        [
            "Telegram Trigger not active: activate the workflow and send any message once.",
            "Wrong sheet / tab name: n8n fails — keep Students_Data, subject, Lessons_Data, student_progress exact.",
            "Gemini or PDFShift API key missing: lesson path stops before email.",
            "Student email empty in Students_Data: Gmail nodes cannot deliver.",
            "Bot token or API keys on GitHub: never commit secrets; credentials stay in n8n only.",
            "Duplicate Telegram error messages can appear on rapid retries.",
        ],
    )

    heading(doc, "25. Conclusion")
    add_p(
        doc,
        "AdaptiveEdu is a complete live loop: identify the student, choose a subject, generate the due lesson with AI, "
        "check understanding with a scored Excel quiz, then either continue or retry until Pass. "
        "Proof for the panel is the bot chat, the lesson email, and the Pass/Fail emails — plus the three Sheets and the n8n canvas.",
    )
    add_p(doc, BOT, size=12, bold=True, color=TEAL, space=2)
    add_p(doc, GITHUB, size=12, bold=True, color=TEAL)
    doc.save(OUT_DOC)
    print(OUT_DOC)


if __name__ == "__main__":
    build_script()
    build_docs()
