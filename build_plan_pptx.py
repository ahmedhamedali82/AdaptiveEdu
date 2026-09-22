#!/usr/bin/env python3
"""AdaptiveEdu — lean project plan (no repeated slides)."""

from pathlib import Path
import shutil

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path("/Users/ahmedhamed/Documents/AI Agent/Technical/Graduation_Project/AdaptiveEdu")
ASSETS = Path("/Users/ahmedhamed/.cursor/projects/Users-ahmedhamed-Documents-AI-Agent/assets")
OUT = ROOT / "AdaptiveEdu_Project_Plan.pptx"
DIAG = ROOT / "diagrams"
DIAG.mkdir(parents=True, exist_ok=True)

W, H = Inches(13.333), Inches(7.5)
NAVY = RGBColor(0x0F, 0x27, 0x44)
NAVY2 = RGBColor(0x1B, 0x3A, 0x61)
TEAL = RGBColor(0x0E, 0x8C, 0x7E)
GOLD = RGBColor(0xE6, 0x9B, 0x2D)
ORANGE = RGBColor(0xE8, 0x78, 0x30)
GREEN = RGBColor(0x2E, 0x8C, 0x5A)
LIGHT = RGBColor(0xF6, 0xF8, 0xFB)
SOFT_TEAL = RGBColor(0xE4, 0xF4, 0xF1)
SOFT_BLUE = RGBColor(0xE6, 0xEE, 0xF8)
SOFT_GOLD = RGBColor(0xFF, 0xF6, 0xE4)
SOFT_GREEN = RGBColor(0xE6, 0xF5, 0xEC)
SOFT_ORANGE = RGBColor(0xFF, 0xF0, 0xE4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1C, 0x24, 0x30)
MUTED = RGBColor(0x5A, 0x66, 0x76)
PAGE = {"n": 0}


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


def add_text(slide, l, t, w, h, text, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
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
    return box


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def chrome(slide, footer):
    add_rect(slide, 0, 0, W, H, WHITE)
    add_rect(slide, 0, 0, Inches(0.12), H, NAVY)
    add_rect(slide, Inches(0.12), Inches(7.22), Inches(13.213), Inches(0.28), LIGHT)
    add_text(slide, Inches(0.35), Inches(7.22), Inches(10), Inches(0.28), footer, size=11, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    PAGE["n"] += 1
    add_text(slide, Inches(11.5), Inches(7.22), Inches(1.5), Inches(0.28), str(PAGE["n"]), size=11, color=MUTED, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def header(slide, kicker, title, footer):
    chrome(slide, footer)
    add_text(slide, Inches(0.4), Inches(0.16), Inches(12.4), Inches(0.26), kicker, size=13, bold=True, color=NAVY)
    add_text(slide, Inches(0.4), Inches(0.40), Inches(12.4), Inches(0.44), title, size=24, bold=True, color=NAVY2)
    add_rect(slide, Inches(0.4), Inches(0.88), Inches(1.4), Inches(0.06), NAVY)


def card(slide, l, t, w, h, fill, title, body, bar=None, tsize=16, bsize=14):
    add_round(slide, l, t, w, h, fill)
    if bar:
        add_rect(slide, l, t, Inches(0.1), h, bar)
    add_text(slide, l + Inches(0.2), t + Inches(0.14), w - Inches(0.35), Inches(0.4), title, size=tsize, bold=True, color=NAVY)
    add_text(slide, l + Inches(0.2), t + Inches(0.55), w - Inches(0.35), h - Inches(0.7), body, size=bsize, color=DARK)


def pic(slide, path, left, top, width):
    slide.shapes.add_picture(str(path), left, top, width=width)


def copy_imgs():
    mapping = {
        "mylesson_overview.png": "personatutor_overview.png",
        "mylesson_n8n_flow.png": "personatutor_n8n_flow.png",
        "personatutor_milestones.png": "personatutor_milestones.png",
    }
    for src_name, dest_name in mapping.items():
        srcp = ASSETS / src_name
        if srcp.exists():
            shutil.copy2(srcp, DIAG / dest_name)


def build():
    PAGE["n"] = 0
    copy_imgs()
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    f = "AdaptiveEdu  ·  Project plan"

    # 1 Cover
    s = blank(prs)
    add_rect(s, 0, 0, W, H, NAVY)
    add_rect(s, Inches(8.6), 0, Inches(4.733), H, NAVY2)
    add_text(s, Inches(0.55), Inches(1.5), Inches(7.6), Inches(0.35), "DEPI GRADUATION PROJECT", size=14, bold=True, color=TEAL)
    add_text(s, Inches(0.55), Inches(2.0), Inches(7.6), Inches(1.2), "AdaptiveEdu", size=44, bold=True, color=WHITE)
    add_text(s, Inches(0.55), Inches(3.4), Inches(7.6), Inches(1.4), "Personal lessons with LLM + n8n.\nNot one fixed book for every student.", size=20, color=WHITE)
    add_text(s, Inches(0.55), Inches(5.2), Inches(7.6), Inches(0.8), "4-week plan  ·  Proposal → Design → Build → Deliver", size=16, color=WHITE)
    for i, line in enumerate(["Problem", "n8n workflow", "Tools & outcomes", "4-week milestones", "Deliverables", "Do this week"]):
        add_text(s, Inches(8.9), Inches(1.8 + i * 0.7), Inches(4.1), Inches(0.55), f"{i+1}  {line}", size=16, color=WHITE)
    PAGE["n"] += 1
    add_text(s, Inches(12.2), Inches(7.15), Inches(0.9), Inches(0.28), "1", size=11, color=WHITE, align=PP_ALIGN.RIGHT)

    # 2 Problem + solution (was slides 2+3+4)
    s = blank(prs)
    header(s, "WHY + WHAT", "Problem and solution", f)
    card(
        s, Inches(0.4), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_ORANGE,
        "Problem",
        "One fixed curriculum for all.\n\nSame pace. Same examples.\n\nWeak students fall behind.\nStrong students get bored.\n\nTeachers cannot rewrite every lesson by hand.",
        bar=ORANGE, tsize=18, bsize=16,
    )
    card(
        s, Inches(6.75), Inches(1.15), Inches(6.15), Inches(5.7), SOFT_GREEN,
        "Solution — AdaptiveEdu",
        "A personalized AI agent.\n\nStudent asks → agent builds content for that student.\n\nInput: prompt + level + goal\nProcess: n8n + LLM + profile\nOutput: personal lesson + quiz\n\nLearning fits the learner.",
        bar=GREEN, tsize=18, bsize=16,
    )

    # 3 Objectives (slim — 4 only, drop overlap with outcomes)
    s = blank(prs)
    header(s, "OBJECTIVES", "What we will achieve", f)
    objs = [
        ("1  Personalize", "Generate content per student request and level."),
        ("2  Automate", "Run the full flow in n8n with little manual work."),
        ("3  Track", "Save each request and output for teacher review."),
        ("4  Deliver", "ZIP, PPT, docs, GitHub, 2–5 min video."),
    ]
    for i, (t, b) in enumerate(objs):
        r, c = divmod(i, 2)
        card(s, Inches(0.4 + c * 6.35), Inches(1.2 + r * 2.85), Inches(6.15), Inches(2.65), LIGHT, t, b, bar=TEAL, tsize=18, bsize=16)

    # 4 How it works = user flow + n8n (was 6+7+8)
    s = blank(prs)
    header(s, "HOW IT WORKS", "Student path + n8n nodes", f)
    pic(s, DIAG / "personatutor_n8n_flow.png", Inches(0.4), Inches(1.05), Inches(12.5))
    # tiny caption strip under image if space - add 6 short node labels as footer cards
    nodes = [
        ("1 Trigger", "Form / Telegram"),
        ("2 Profile", "Google Sheet"),
        ("3 LLM", "Generate lesson"),
        ("4 Format", "Code / Set"),
        ("5 Log", "Save run"),
        ("6 Send", "Email / Telegram"),
    ]
    for i, (t, b) in enumerate(nodes):
        x = 0.4 + i * 2.12
        add_round(s, Inches(x), Inches(5.85), Inches(2.02), Inches(1.15), SOFT_TEAL if i % 2 == 0 else LIGHT)
        add_text(s, Inches(x + 0.08), Inches(5.95), Inches(1.86), Inches(0.4), t, size=13, bold=True, color=NAVY)
        add_text(s, Inches(x + 0.08), Inches(6.35), Inches(1.86), Inches(0.5), b, size=12, color=DARK)

    # 5 Prompt (unique — keep)
    s = blank(prs)
    header(s, "PROMPT TEMPLATE", "Same structure. Different student data.", f)
    add_round(s, Inches(0.4), Inches(1.15), Inches(12.5), Inches(5.7), SOFT_BLUE)
    add_text(
        s, Inches(0.7), Inches(1.4), Inches(12.0), Inches(5.2),
        "Role: You are a school tutor in Egypt.\n"
        "Student: {{name}}  ·  Level: {{level}}  ·  Subject: {{subject}}\n"
        "Request: {{student_prompt}}\n\n"
        "Task: Create a short personal lesson.\n"
        "Format:\n"
        "1) Title\n"
        "2) 5 key points (simple English)\n"
        "3) One Egypt school example\n"
        "4) 3 practice questions\n"
        "5) One tip for next study session\n\n"
        "Rules: No fake facts. If unsure, say so. Keep under 350 words.",
        size=18, color=DARK,
    )

    # 6 Tools + outcomes (was 10+11)
    s = blank(prs)
    header(s, "TOOLS & OUTCOMES", "Stack + what success looks like", f)
    tools = [
        ("n8n", "Workflow: trigger → LLM → send"),
        ("LLM", "ChatGPT / Gemini / Groq"),
        ("Sheets", "Profiles + activity log"),
        ("Telegram / Gmail", "Ask and deliver"),
    ]
    for i, (t, b) in enumerate(tools):
        card(s, Inches(0.4 + i * 3.2), Inches(1.15), Inches(3.05), Inches(2.5), SOFT_TEAL, t, b, bar=TEAL, tsize=15, bsize=14)
    outs = [
        ("Working agent", "Ask → personal lesson arrives."),
        ("Workflow JSON", "Importable n8n + AI node."),
        ("Teacher log", "Sheet of requests/outputs."),
        ("DEPI package", "ZIP, PPT, docs, GitHub, video."),
    ]
    for i, (t, b) in enumerate(outs):
        card(s, Inches(0.4 + i * 3.2), Inches(3.95), Inches(3.05), Inches(2.9), SOFT_GREEN, t, b, bar=GREEN, tsize=15, bsize=14)

    # 7 Milestones 4 weeks (was 12+13)
    s = blank(prs)
    header(s, "MILESTONES", "4-week delivery", f)
    miles = [
        ("Week 1", "Proposal + GitHub", "Problem, objectives, tools, outcomes.\nTeam roles. Fixed GitHub URL."),
        ("Week 2", "Design", "n8n canvas. Prompt templates.\nSheet columns. User flow."),
        ("Week 3", "Build + test", "Trigger → profile → LLM → send.\nTest 3 student personas."),
        ("Week 4", "Docs + delivery", "Docs, final PPT, 2–5 min video.\nZIP + defense prep."),
    ]
    for i, (w, t, b) in enumerate(miles):
        card(s, Inches(0.4 + i * 3.2), Inches(1.2), Inches(3.05), Inches(5.6), [SOFT_TEAL, SOFT_BLUE, SOFT_GOLD, SOFT_GREEN][i], f"{w}\n{t}", b, bar=[TEAL, NAVY2, ORANGE, GREEN][i], tsize=18, bsize=15)

    # 8 Deliverables + roles (was 14+15)
    s = blank(prs)
    header(s, "DELIVERABLES & ROLES", "What we submit  ·  who does what", f)
    dels = [
        ("Source ZIP", "n8n JSON + prompts"),
        ("Final PPT", "DEPI panel deck"),
        ("Full docs", "Proposal + user guide"),
        ("GitHub", "Create early. Never change URL."),
        ("Video 2–5 min", "Problem → demo → result"),
        ("15-min defense", "Each member’s role clear"),
    ]
    for i, (t, b) in enumerate(dels):
        r, c = divmod(i, 3)
        card(s, Inches(0.4 + c * 4.22), Inches(1.15 + r * 1.85), Inches(4.05), Inches(1.7), LIGHT, t, b, bar=TEAL, tsize=15, bsize=13)
    roles = "Leader: deadlines + GitHub  ·  n8n: workflow  ·  Prompts: templates  ·  Docs/video: PPT + film  ·  QA: tests + sheet log"
    add_round(s, Inches(0.4), Inches(5.1), Inches(12.5), Inches(1.8), SOFT_BLUE)
    add_text(s, Inches(0.7), Inches(5.35), Inches(12.0), Inches(0.4), "Team roles (fill names later)", size=16, bold=True, color=NAVY)
    add_text(s, Inches(0.7), Inches(5.85), Inches(12.0), Inches(0.8), roles, size=15, color=DARK)

    # 9 Do this week (was 16+17)
    s = blank(prs)
    header(s, "DO THIS WEEK", "Then we write the proposal document", f)
    now = [
        ("1", "Confirm name", "AdaptiveEdu"),
        ("2", "Register team", "Form before deadline"),
        ("3", "Create GitHub", "Empty repo. Link stays forever."),
        ("4", "Draft proposal", "Problem · objectives · tools · outcomes"),
        ("5", "Sketch n8n", "6 nodes on paper. Build in Week 3."),
    ]
    for i, (n, t, b) in enumerate(now):
        y = 1.2 + i * 1.1
        add_round(s, Inches(0.4), Inches(y), Inches(12.5), Inches(1.0), SOFT_TEAL if i % 2 == 0 else LIGHT)
        add_round(s, Inches(0.55), Inches(y + 0.22), Inches(0.55), Inches(0.55), TEAL)
        add_text(s, Inches(0.55), Inches(y + 0.22), Inches(0.55), Inches(0.55), n, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(1.35), Inches(y + 0.15), Inches(4.5), Inches(0.7), t, size=18, bold=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_text(s, Inches(6.0), Inches(y + 0.15), Inches(6.6), Inches(0.7), b, size=16, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

    prs.save(OUT)
    print(OUT)
    print("slides", len(prs.slides))


if __name__ == "__main__":
    build()
