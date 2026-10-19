#!/usr/bin/env python3
"""AdaptiveEdu project plan — live Telegram bot + Gemini lessons + 3 Sheets."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path("/Users/ahmedhamed/Documents/AI Agent/Technical/Graduation_Project/AdaptiveEdu")
OUT = ROOT / "AdaptiveEdu_Project_Plan.pptx"
DIAG = ROOT / "diagrams"
DEMO = DIAG / "demo"

W, H = Inches(13.333), Inches(7.5)
NAVY = RGBColor(0x0F, 0x27, 0x44)
NAVY2 = RGBColor(0x1B, 0x3A, 0x61)
TEAL = RGBColor(0x0E, 0x8C, 0x7E)
GOLD = RGBColor(0xE6, 0x9B, 0x2D)
ORANGE = RGBColor(0xE8, 0x78, 0x30)
GREEN = RGBColor(0x2E, 0x8C, 0x5A)
RED = RGBColor(0xB9, 0x1C, 0x1C)
LIGHT = RGBColor(0xF6, 0xF8, 0xFB)
SOFT_TEAL = RGBColor(0xE4, 0xF4, 0xF1)
SOFT_BLUE = RGBColor(0xE6, 0xEE, 0xF8)
SOFT_GOLD = RGBColor(0xFF, 0xF6, 0xE4)
SOFT_GREEN = RGBColor(0xE6, 0xF5, 0xEC)
SOFT_ORANGE = RGBColor(0xFF, 0xF0, 0xE4)
SOFT_RED = RGBColor(0xFF, 0xEC, 0xE9)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1C, 0x24, 0x30)
MUTED = RGBColor(0x5A, 0x66, 0x76)
PAGE = {"n": 0}

GITHUB = "https://github.com/ahmedhamedali82/AdaptiveEdu"
BOT = "https://t.me/AdaptiveEdubot"
SHEET_STU = "https://docs.google.com/spreadsheets/d/1ddIwenYjiwKryP7Pogxe-7qsjKnTmtfCtAlDoplSn5Q"
SHEET_LES = "https://docs.google.com/spreadsheets/d/1akeM65odle57kxtR7M5LqFMnxCsHX3M1cJKKltExFEg"
SHEET_PRG = "https://docs.google.com/spreadsheets/d/1II6ckyiKFucKsC38PC_0a9tH1h8H0JZe3qaclLxCExY"


def rgb_fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_rect(slide, l, t, w, h, color):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    rgb_fill(sh, color)
    return sh


def add_round(slide, l, t, w, h, color, adj=0.08):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    rgb_fill(sh, color)
    try:
        sh.adjustments[0] = adj
    except Exception:
        pass
    return sh


def add_text(slide, l, t, w, h, text, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, url=None):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    if url:
        run.hyperlink.address = url
    return box


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def chrome(slide, footer):
    add_rect(slide, 0, 0, W, H, WHITE)
    add_rect(slide, 0, 0, Inches(0.12), H, NAVY)
    add_rect(slide, Inches(0.12), Inches(7.22), Inches(13.213), Inches(0.28), LIGHT)
    add_text(slide, Inches(0.35), Inches(7.22), Inches(10.5), Inches(0.28), footer, size=11, color=MUTED, anchor=MSO_ANCHOR.MIDDLE, url=GITHUB)
    PAGE["n"] += 1
    add_text(slide, Inches(11.5), Inches(7.22), Inches(1.5), Inches(0.28), str(PAGE["n"]), size=11, color=MUTED, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def header(slide, kicker, title, footer):
    chrome(slide, footer)
    add_text(slide, Inches(0.4), Inches(0.16), Inches(12.4), Inches(0.26), kicker, size=13, bold=True, color=TEAL)
    add_text(slide, Inches(0.4), Inches(0.40), Inches(12.4), Inches(0.44), title, size=24, bold=True, color=NAVY2)
    add_rect(slide, Inches(0.4), Inches(0.88), Inches(1.4), Inches(0.06), NAVY)


def card(slide, l, t, w, h, fill, title, body, bar=None, tsize=16, bsize=14):
    add_round(slide, l, t, w, h, fill)
    if bar:
        add_rect(slide, l, t, Inches(0.1), h, bar)
    add_text(slide, l + Inches(0.2), t + Inches(0.14), w - Inches(0.35), Inches(0.4), title, size=tsize, bold=True, color=NAVY)
    add_text(slide, l + Inches(0.2), t + Inches(0.55), w - Inches(0.35), h - Inches(0.7), body, size=bsize, color=DARK)


def pic(slide, path, left, top, width=None, height=None):
    p = Path(path)
    if not p.exists():
        return
    kwargs = {}
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    slide.shapes.add_picture(str(p), left, top, **kwargs)


def build():
    PAGE["n"] = 0
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    f = GITHUB

    # 1 Cover
    s = blank(prs)
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, Inches(8.6), 0, Inches(4.733), H, NAVY2)
    add_text(s, Inches(0.55), Inches(1.2), Inches(7.6), Inches(0.32), "DEPI GRADUATION PROJECT", size=14, bold=True, color=TEAL)
    add_text(s, Inches(0.55), Inches(1.7), Inches(7.6), Inches(0.9), "AdaptiveEdu", size=44, bold=True, color=WHITE)
    add_text(
        s, Inches(0.55), Inches(2.8), Inches(7.6), Inches(1.8),
        "Telegram study bot for 1st Secondary.\nRegister → choose subject → AI lesson PDF\n→ submit Excel quiz → Pass / Fail by email.",
        size=18, color=WHITE,
    )
    add_text(s, Inches(0.55), Inches(4.85), Inches(7.6), Inches(0.35), "n8n  ·  Gemini  ·  3 Google Sheets  ·  Gmail  ·  Drive", size=15, color=WHITE)
    add_text(s, Inches(0.55), Inches(5.35), Inches(7.6), Inches(0.35), BOT, size=16, bold=True, color=TEAL, url=BOT)
    add_text(s, Inches(0.55), Inches(5.85), Inches(7.6), Inches(0.35), GITHUB, size=14, color=TEAL, url=GITHUB)
    toc = ["Problem", "Student journey", "3 Sheets", "n8n steps", "Live canvas", "Demo emails", "Pass / Fail", "Deliverables"]
    for i, line in enumerate(toc):
        add_text(s, Inches(8.9), Inches(1.45 + i * 0.6), Inches(4.1), Inches(0.5), f"{i+1}  {line}", size=16, color=WHITE)
    PAGE["n"] += 1
    add_text(s, Inches(12.2), Inches(7.15), Inches(0.9), Inches(0.28), "1", size=11, color=WHITE, align=PP_ALIGN.RIGHT)

    # 2 Problem + solution
    s = blank(prs)
    header(s, "WHY THIS PROJECT", "Problem and solution", f)
    card(
        s, Inches(0.4), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_ORANGE,
        "Problem",
        "One class. One pace. One book.\n\nA student who fails still moves on.\nA teacher cannot follow every student by hand.\n\nWe need: finish this lesson before the next one — with proof.",
        bar=ORANGE, tsize=20, bsize=17,
    )
    card(
        s, Inches(6.75), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_GREEN,
        "Solution — AdaptiveEdu",
        "A live Telegram bot.\n\n1) Student sends ID (ST005).\n2) Chooses a subject code.\n3) Gemini builds the next lesson as PDF + 10 MCQs.\n4) Student uploads Excel answers.\n5) Score saved. Email: Pass or Fail.\n\nBot: t.me/AdaptiveEdubot",
        bar=GREEN, tsize=20, bsize=16,
    )

    # 3 Student journey — clear steps
    s = blank(prs)
    header(s, "STUDENT JOURNEY", "Eight clear steps (demo path)", f)
    steps = [
        ("1", "Open bot", "t.me/AdaptiveEdubot"),
        ("2", "Send ID", "ST005"),
        ("3", "Pick subject", "Send 4 = Philosophy"),
        ("4", "Get lesson", "PDF attached by email"),
        ("5", "Study", "Open PDF attachment"),
        ("6", "Answer file", "ST005_4_1.xlsx"),
        ("7", "Upload Excel", "On Telegram"),
        ("8", "Result email", "Pass (>6) or Fail"),
    ]
    for i, (n, t, b) in enumerate(steps):
        r, c = divmod(i, 4)
        x = 0.4 + c * 3.2
        y = 1.2 + r * 2.95
        add_round(s, Inches(x), Inches(y), Inches(3.05), Inches(2.7), [SOFT_TEAL, SOFT_BLUE, SOFT_GOLD, SOFT_GREEN][c])
        add_round(s, Inches(x + 1.1), Inches(y + 0.25), Inches(0.7), Inches(0.7), TEAL)
        add_text(s, Inches(x + 1.1), Inches(y + 0.25), Inches(0.7), Inches(0.7), n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(x + 0.15), Inches(y + 1.15), Inches(2.75), Inches(0.5), t, size=17, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(s, Inches(x + 0.15), Inches(y + 1.7), Inches(2.75), Inches(0.7), b, size=14, color=DARK, align=PP_ALIGN.CENTER)

    # 4 Three sheets
    s = blank(prs)
    header(s, "DATA", "Three Google Sheets the bot uses", f)
    card(
        s, Inches(0.4), Inches(1.15), Inches(4.05), Inches(5.7), SOFT_BLUE,
        "1  Students_Data",
        "Who is the student?\n\nstudent_id\nstudent_name\ntelegram_chat_id\nemail\nstatus (waiting flag)\n\nFirst login saves chat ID.",
        bar=NAVY2, tsize=18, bsize=15,
    )
    card(
        s, Inches(4.65), Inches(1.15), Inches(4.05), Inches(5.7), SOFT_TEAL,
        "2  Lessons_Data",
        "What can we teach?\n\nTab subject:\n1–8 subject codes\n\nTab Lessons_Data:\nlesson_code + lesson_name\nper subject\n\nNext lesson = first not passed.",
        bar=TEAL, tsize=18, bsize=15,
    )
    card(
        s, Inches(8.9), Inches(1.15), Inches(4.05), Inches(5.7), SOFT_GOLD,
        "3  student_progress",
        "History log.\n\nstudent + subject + lesson\ndate, grade, succeed\nFILE_NAME\nANS_1 … ANS_10\n\nAppend / update each attempt.",
        bar=ORANGE, tsize=18, bsize=15,
    )

    # 5 Sheet links
    s = blank(prs)
    header(s, "SHEET LINKS", "Open these in the live demo", f)
    card(s, Inches(0.4), Inches(1.15), Inches(12.5), Inches(1.7), SOFT_BLUE, "Students_Data", SHEET_STU, bar=NAVY2, tsize=18, bsize=14)
    card(s, Inches(0.4), Inches(3.05), Inches(12.5), Inches(1.7), SOFT_TEAL, "Lessons_Data (subject + lessons)", SHEET_LES, bar=TEAL, tsize=18, bsize=14)
    card(s, Inches(0.4), Inches(4.95), Inches(12.5), Inches(1.7), SOFT_GOLD, "student_progress", SHEET_PRG, bar=ORANGE, tsize=18, bsize=14)

    # 6 Workflow steps — register path
    s = blank(prs)
    header(s, "N8N STEPS  ·  A", "Register path (first time)", f)
    flow_a = [
        ("1 Telegram Trigger", "New message arrives."),
        ("2 Get row by chat ID", "Look up Students_Data."),
        ("3 status ≠ waiting", "Treat as register / menu."),
        ("4 Search student ID", "Match text like ST005."),
        ("5 Update chat ID", "Save Telegram id once."),
        ("6 Show subjects", "List 1–8 + upload tip."),
        ("7 Set waiting=1", "Next message = subject or file."),
    ]
    for i, (t, b) in enumerate(flow_a):
        y = 1.1 + i * 0.82
        add_round(s, Inches(0.4), Inches(y), Inches(12.5), Inches(0.72), SOFT_TEAL if i % 2 == 0 else LIGHT)
        add_round(s, Inches(0.55), Inches(y + 0.12), Inches(0.48), Inches(0.48), TEAL)
        add_text(s, Inches(0.55), Inches(y + 0.12), Inches(0.48), Inches(0.48), str(i + 1), size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.25), Inches(y + 0.1), Inches(4.5), Inches(0.52), t, size=16, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(5.9), Inches(y + 0.1), Inches(6.7), Inches(0.52), b, size=15, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

    # 7 Workflow steps — lesson path
    s = blank(prs)
    header(s, "N8N STEPS  ·  B", "Lesson path (subject code)", f)
    flow_b = [
        ("1 Subject check", "Text equals SUBJECT_CODE."),
        ("2 Read progress", "Find last passed lesson."),
        ("3 Pick next lesson", "Code node get_file."),
        ("4 Block if pending", "Old lesson unanswered → error."),
        ("5 AI Agent (Gemini)", "Markdown lesson + 10 answers."),
        ("6 HTML → PDFShift", "HTTP Request returns binary PDF."),
        ("7 Gmail + Telegram", "Email PDF attachment + Excel name; Telegram notice."),
        ("8 Append progress", "Open lesson row ready for grading."),
    ]
    for i, (t, b) in enumerate(flow_b):
        r, c = divmod(i, 4)
        card(s, Inches(0.35 + c * 3.2), Inches(1.15 + r * 2.95), Inches(3.05), Inches(2.75), LIGHT if i % 2 == 0 else SOFT_TEAL, t, b, bar=TEAL, tsize=15, bsize=14)

    # 8 Workflow steps — grade path
    s = blank(prs)
    header(s, "N8N STEPS  ·  C", "Grade path (Excel upload)", f)
    flow_c = [
        ("1 Document?", "If message has a file."),
        ("2 Name check", "Must match ST###_sub_les.xlsx"),
        ("3 Drive extract", "Upload → download → parse row."),
        ("4 Compare ANS_1..10", "Capital letters vs model answers."),
        ("5 Score", "grade = correct count / 10"),
        ("6 Pass rule", "Pass if grade > 6  else Fail"),
        ("7 Update sheet", "Write grade + succeed flag."),
        ("8 Email + Telegram", "Result mail + restudy tip if Fail."),
    ]
    for i, (t, b) in enumerate(flow_c):
        r, c = divmod(i, 4)
        card(s, Inches(0.35 + c * 3.2), Inches(1.15 + r * 2.95), Inches(3.05), Inches(2.75), SOFT_GOLD if i < 4 else SOFT_GREEN, t, b, bar=ORANGE if i < 4 else GREEN, tsize=15, bsize=14)

    # 9 Live n8n canvas
    s = blank(prs)
    header(s, "N8N CANVAS", "Live workflow ADAPTIVE_14 (46 nodes)", f)
    pic(s, DEMO / "workflow_n8n.jpeg", Inches(0.25), Inches(1.05), width=Inches(12.8))
    add_text(
        s, Inches(0.4), Inches(6.45), Inches(12.5), Inches(0.5),
        "Three message types only:  student ID  ·  subject code  ·  answers Excel file",
        size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
    )

    # 10 Tools
    s = blank(prs)
    header(s, "TOOLS", "What each tool does", f)
    tools = [
        ("Telegram", "Student screen. Bot link on cover."),
        ("n8n", "Orchestrates all paths."),
        ("Google Sheets ×3", "Students, lessons, progress."),
        ("Gemini AI Agent", "Writes lesson + model answers."),
        ("PDFShift", "HTML → binary PDF for email."),
        ("Gmail + Drive", "Attach PDF + store answer files."),
    ]
    for i, (t, b) in enumerate(tools):
        r, c = divmod(i, 3)
        card(s, Inches(0.4 + c * 4.22), Inches(1.15 + r * 2.9), Inches(4.05), Inches(2.7), [SOFT_TEAL, SOFT_BLUE, SOFT_GOLD][c], t, b, bar=TEAL, tsize=18, bsize=16)

    # 11 Quiz Excel style
    s = blank(prs)
    header(s, "QUIZ FILE", "Exact Excel style the bot accepts", f)
    pic(s, DEMO / "quiz_excel_style.png", Inches(0.55), Inches(1.15), width=Inches(12.2))
    add_text(
        s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.5),
        "Example for Ahmed (ST005), Philosophy (4), lesson 1  →  file name ST005_4_1.xlsx",
        size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
    )

    # 12 Bot screenshot
    s = blank(prs)
    header(s, "LIVE DEMO  ·  TELEGRAM", "Bot messages (subjects + validation)", f)
    bot_img = DEMO / "bot_chat.jpeg" if (DEMO / "bot_chat.jpeg").exists() else DEMO / "bot_errors.jpg"
    pic(s, bot_img, Inches(0.35), Inches(1.05), height=Inches(5.9))
    card(
        s, Inches(5.15), Inches(1.15), Inches(7.75), Inches(2.5), SOFT_TEAL,
        "What you should see",
        "After a valid ID, the bot lists subjects 1–8 and offers to upload an answers file.\nSend a subject number (example: 4) to open the next lesson.",
        bar=TEAL, tsize=18, bsize=15,
    )
    card(
        s, Inches(5.15), Inches(3.9), Inches(7.75), Inches(2.85), SOFT_ORANGE,
        "Validation messages",
        "Wrong subject → “كود المادة غير موجود…”\nWrong / unknown student ID → “كود الطالب غير صحيح…”\nThese prove the Sheets checks are live.",
        bar=ORANGE, tsize=18, bsize=15,
    )

    # 13 Telegram + email lesson delivery
    s = blank(prs)
    header(s, "LIVE DEMO  ·  LESSON DELIVERY", "Telegram notice + Gmail with PDF attachment", f)
    pic(s, DEMO / "msg_lesson_sent.png", Inches(0.35), Inches(1.1), width=Inches(5.9))
    pic(s, DEMO / "email_lesson.png", Inches(6.5), Inches(1.1), width=Inches(6.4))
    add_text(
        s, Inches(0.4), Inches(6.4), Inches(12.5), Inches(0.5),
        "Real email example: Philosophy & Logic → Thinking Principles → PDF attachment → reply with ST005_4_1.xlsx",
        size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
    )

    # 14 Pass / Fail emails
    s = blank(prs)
    header(s, "LIVE DEMO  ·  RESULTS", "Same student — Fail then Pass emails", f)
    pic(s, DEMO / "email_fail.png", Inches(0.4), Inches(1.15), width=Inches(6.1))
    pic(s, DEMO / "email_pass.png", Inches(6.8), Inches(1.15), width=Inches(6.1))
    add_text(s, Inches(0.4), Inches(5.55), Inches(6.1), Inches(0.4), "Fail example: 2 / 10", size=16, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_text(s, Inches(6.8), Inches(5.55), Inches(6.1), Inches(0.4), "Pass example: 9 / 10", size=16, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(
        s, Inches(0.5), Inches(6.15), Inches(12.3), Inches(0.7),
        "Rule in n8n: Pass if grade > 6. Fail → restudy same lesson and upload again. Pass → next lesson_code opens.",
        size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
    )

    # 15 Rule slide
    s = blank(prs)
    header(s, "RULE", "Pass, fail, and what happens next", f)
    card(
        s, Inches(0.4), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_GREEN,
        "Pass  ·  more than 6 / 10",
        "student_progress: succeed = 1\ngrade saved with ANS_1…ANS_10\n\nEmail: “you are Pass”\nTelegram: grading done.\n\nNext subject request opens lesson_code + 1.",
        bar=GREEN, tsize=20, bsize=17,
    )
    card(
        s, Inches(6.75), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_RED,
        "Fail  ·  6 or less / 10",
        "student_progress: succeed = 0\n\nEmail: “you are Fail”\nTelegram: restudy and resend answers.\n\nSame lesson stays open until Pass.\nWrong file name is refused.",
        bar=RED, tsize=20, bsize=17,
    )

    # 16 Message cards
    s = blank(prs)
    header(s, "BOT + EMAIL TEXT", "Key messages the panel can read aloud", f)
    pic(s, DEMO / "msg_welcome.png", Inches(0.3), Inches(1.1), width=Inches(6.2))
    pic(s, DEMO / "msg_graded.png", Inches(6.8), Inches(1.1), width=Inches(6.2))
    add_text(
        s, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.5),
        "Left: welcome + subject list after register.   Right: after Excel upload (before the result email).",
        size=14, color=MUTED, align=PP_ALIGN.CENTER,
    )

    # 17 Deliverables
    s = blank(prs)
    header(s, "SUBMIT", "What the panel should see", f)
    dels = [
        ("Working bot", BOT.replace("https://", "")),
        ("n8n JSON", "AdaptiveEdu.json (ADAPTIVE_14)"),
        ("3 Sheets", "Students · Lessons · Progress"),
        ("This PPT", "Journey + n8n steps + live shots"),
        ("Docs + video", "Full write-up + 2–5 min demo"),
        ("GitHub", GITHUB.replace("https://", "")),
    ]
    for i, (t, b) in enumerate(dels):
        r, c = divmod(i, 3)
        card(s, Inches(0.4 + c * 4.22), Inches(1.15 + r * 2.9), Inches(4.05), Inches(2.7), LIGHT, t, b, bar=TEAL, tsize=18, bsize=15)

    # 18 Demo script
    s = blank(prs)
    header(s, "LIVE DEMO SCRIPT", "Say this. Do this. ~3 minutes.", f)
    demo = [
        ("1", "Open Telegram bot", "Send ST005"),
        ("2", "Show subject list", "Send 4 (Philosophy)"),
        ("3", "Open Gmail", "PDF attachment + file name ST005_4_1.xlsx"),
        ("4", "Show Excel row", "10 capital letters, no header"),
        ("5", "Upload file on Telegram", "Bot confirms grading"),
        ("6", "Show Pass or Fail email", "Then open student_progress row"),
    ]
    for i, (n, t, b) in enumerate(demo):
        y = 1.15 + i * 0.95
        add_round(s, Inches(0.4), Inches(y), Inches(12.5), Inches(0.85), SOFT_TEAL if i % 2 == 0 else LIGHT)
        add_round(s, Inches(0.55), Inches(y + 0.15), Inches(0.55), Inches(0.55), TEAL)
        add_text(s, Inches(0.55), Inches(y + 0.15), Inches(0.55), Inches(0.55), n, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.35), Inches(y + 0.12), Inches(4.5), Inches(0.6), t, size=17, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(6.0), Inches(y + 0.12), Inches(6.6), Inches(0.6), b, size=16, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

    prs.save(OUT)
    print(OUT)
    print("slides", len(prs.slides))


if __name__ == "__main__":
    build()
